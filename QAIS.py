# -*- coding: utf-8 -*-
"""
QAIS - Intelligent AIS Vessel Tracking and Visualisation

2025-11-01 (v1.0.1)
- Copyright (C) 2025 Spatial Exploration Pty Ltd
- QAIS reads, decodes and displays live AIS targets on the QGIS map layout. 
- User can target and track indivdual vessels and export all AIS messages for further analysis. 
- This plugin is not to be used for navigational purposes. 
- This plugin requires an external AIS receiver to operate. These can be purchased from www.wegmatt.com/index.html. 
- Proudly designed and built in Western Australia by Spatial Exploration www.spatial-exploration.com
- This program is free software; you can redistribute it and/or modify
- it under the terms of the GNU General Public License as published by
- the Free Software Foundation; either version 3 of the License, or
- (at your option) any later version.

"""

import os
import time
import re

from qgis.PyQt.QtCore import QThread, pyqtSignal, QVariant, Qt
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import QAction

from qgis.core import (
    QgsProject, QgsVectorLayer, QgsField, QgsFeature, QgsGeometry, QgsPointXY,
    QgsMarkerSymbol, QgsRuleBasedRenderer, QgsFeatureRequest,
    QgsGeometryGeneratorSymbolLayer, QgsFillSymbol, QgsUnitTypes,
    QgsSvgMarkerSymbolLayer
)

from .QAIS_dockwidget import QAISDockWidget

# Serial & AIS decode
import serial
import serial.tools.list_ports
try:
    from pyais import decode as ais_decode
    _PY_AIS_OK = True
except Exception as _e:
    _PY_AIS_OK = False
    _PY_AIS_ERR = str(_e)

# Resources (qrc compiled optional)
try:
    import resources_rc  # noqa: F401
except Exception:
    pass


# -------------------- Connect to AIS Receiver via COM Port --------------------
class AISReaderThread(QThread):
    new_message = pyqtSignal(dict)   # {'data': dict, 'ts': float}
    status_msg = pyqtSignal(str)
    error_msg = pyqtSignal(str)

    def __init__(self, port, baud=38400, parent=None):
        super().__init__(parent)
        self._port = port
        self._baud = baud
        self._stop = False

    def run(self):
        try:
            self.status_msg.emit(f"Opening {self._port} @ {self._baud}…")
            with serial.Serial(self._port, self._baud, timeout=1) as ser:
                self.status_msg.emit("Connected. Listening for AIS…")
                while not self._stop:
                    line = ser.readline()
                    if not line:
                        continue
                    nmea = line.decode("ascii", errors="ignore").strip()
                    if not nmea.startswith(("!AIVDM", "!AIVDO")):
                        continue
                    try:
                        msg = ais_decode(nmea) if _PY_AIS_OK else None
                        if msg is None:
                            continue
                        data = msg.asdict()
                        try:
                            data["_RAW_NMEA"] = nmea
                        except Exception:
                            pass
                        self.new_message.emit({"data": data, "ts": time.time()})
                    except Exception as e:
                        self.error_msg.emit(f"Decode error: {e}")
        except Exception as e:
            self.error_msg.emit(f"Serial error: {e}")

    def stop(self):
        self._stop = True


# -------------------- Main Plugin --------------------
class QAISPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.action = None
        self.dock = None
        self.reader = None
        self.ais = None  # working layer
        self.tracked_mmsi = None  # currently selected/target MMSI
        self._connected_selection = False

    
    def initGui(self):
        self.action = QAction(self._load_plugin_icon(), "QAIS", self.iface.mainWindow())
        self.action.triggered.connect(self.show_dock)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu("&QAIS", self.action)

    def unload(self):
        try:
            if self.reader and self.reader.isRunning():
                self.reader.stop()
                self.reader.wait(1000)
        except Exception:
            pass
        try:
            if self.dock:
                self.iface.removeDockWidget(self.dock)
        except Exception:
            pass
        try:
            if self.action:
                self.iface.removeToolBarIcon(self.action)
                self.iface.removePluginMenu("&QAIS", self.action)
        except Exception:
            pass

    def show_dock(self):
        if not self.dock:
            self.dock = QAISDockWidget()
            self.dock.setWindowTitle("QAIS")

            # Serial UI hookup
            if hasattr(self.dock, "cmbBaud"):
                self.dock.cmbBaud.clear()
                self.dock.cmbBaud.addItems(["38400", "9600", "115200"])
                self.dock.cmbBaud.setCurrentText("38400")
            if hasattr(self.dock, "btnRefreshPorts"):
                self.dock.btnRefreshPorts.clicked.connect(self.refresh_ports)
            if hasattr(self.dock, "btnStartStop"):
                self.dock.btnStartStop.clicked.connect(self.toggle_stream)

            if hasattr(self.dock, "TargetVessel_Checkbox"):
                self.dock.TargetVessel_Checkbox.stateChanged.connect(self.toggle_tracking)

            self.refresh_ports()
            self.ensure_layer_and_style()

            # Connect selection listener to the AIS layer itself
            if self.ais and not self._connected_selection:
                try:
                    self.ais.selectionChanged.connect(self.on_selection_changed)
                    self._connected_selection = True
                except Exception:
                    pass

            self.iface.addDockWidget(Qt.LeftDockWidgetArea, self.dock)

        self.dock.show()

    # ---------- COM Port Controls ----------
    def toggle_stream(self):
        if self.reader and self.reader.isRunning():
            self._stop_reader()
            return

        port = getattr(self.dock, "cmbPort", None)
        baud = getattr(self.dock, "cmbBaud", None)
        port_name = port.currentText() if port else None
        try:
            baud_rate = int(baud.currentText()) if baud else 38400
        except Exception:
            baud_rate = 38400

        if not port_name:
            self._set_status("Select a port first")
            return

        self.reader = AISReaderThread(port_name, baud_rate, parent=self.dock)
        self.reader.new_message.connect(self.on_message)
        self.reader.status_msg.connect(self._set_status)
        self.reader.error_msg.connect(self._set_status)
        self.reader.start()
        self._set_status(f"Started {port_name} @ {baud_rate}")
        btn = getattr(self.dock, "btnStartStop", None)
        if btn:
            btn.setText("Stop")

    def _stop_reader(self):
        try:
            if self.reader and self.reader.isRunning():
                self.reader.stop()
                self.reader.wait(1000)
        except Exception:
            pass
        self.reader = None
        self._set_status("Stopped")
        btn = getattr(self.dock, "btnStartStop", None)
        if btn:
            btn.setText("Start")

    def _set_status(self, txt):
        lbl = getattr(self.dock, "lblStatus", None)
        if lbl:
            if "Decode error" in str(txt):
                txt = "Awaiting next AIS message"
            lbl.setText(txt)




    def refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        if hasattr(self.dock, "cmbPort"):
            self.dock.cmbPort.clear()
            self.dock.cmbPort.addItems(ports)
        self._set_status("Select a port and click Start" if ports else "No serial ports found")

    # ---------- AIS Data properties and styling ----------
    def ensure_layer_and_style(self):
        layer = self._get_layer_by_name("AIS Data")
        if layer is None:
            layer = QgsVectorLayer("Point?crs=EPSG:4326", "AIS Data", "memory")
            QgsProject.instance().addMapLayer(layer)
        self.ais = layer
        self._ensure_base_schema(self.ais)
        self._apply_renderer_latest_flag(self.ais)
        self.ais.triggerRepaint()

    def _ensure_base_schema(self, layer):
        pr = layer.dataProvider()
        fields = layer.fields()

        def missing(name: str) -> bool:
            return fields.indexFromName(name) < 0

        need = []
        base = [
            ("MMSI", QVariant.LongLong),
            ("NAME", QVariant.String),
            ("SOG", QVariant.Double),
            ("COG", QVariant.Double),
            ("HDG", QVariant.Double),
            ("LAST_TS", QVariant.Double),
            ("IS_LATEST", QVariant.Int),
            ("IS_TRACKING", QVariant.Int),  # NEW
            ("LAT", QVariant.Double),
            ("LON", QVariant.Double),
            ("DIM_A", QVariant.Double),
            ("DIM_B", QVariant.Double),
            ("DIM_C", QVariant.Double),
            ("DIM_D", QVariant.Double),
            ("LENGTH_M", QVariant.Double),
            ("WIDTH_M", QVariant.Double),
        ]
        for name, typ in base:
            if missing(name):
                f = QgsField(name, type=typ)
                need.append(f)
        if need:
            pr.addAttributes(need)
            layer.updateFields()

    def _apply_renderer_latest_flag(self, layer):
        """
        Apply style from external QML file instead of manual symbol logic.
        This keeps the full dock functionality intact and relies on QGIS
        for rotation, fill, and rule hierarchy.
        """
        qml_path = os.path.join(os.path.dirname(__file__), "resources", "QAIS_Style.qml")
        if os.path.exists(qml_path):
            try:
                layer.loadNamedStyle(qml_path)
                layer.triggerRepaint()
            except Exception as e:
                from qgis.PyQt.QtWidgets import QMessageBox
                QMessageBox.warning(None, "QAIS Style Load Error", f"Failed to apply style: {e}")
        else:
            from qgis.PyQt.QtWidgets import QMessageBox
            QMessageBox.warning(None, "QAIS Style Load Error", "QAIS_Style.qml not found in resources folder.")

    # ---------- AIS Message handling ----------
    def on_message(self, payload: dict):
        data = payload.get("data") or {}
        ts = payload.get("ts", time.time())

        self._check_tracking_timeout()

        mmsi = self._to_int(self._first(data, "mmsi"))
        if mmsi is None:
            return

        lon = self._to_float(self._first(data, "lon", "longitude"))
        lat = self._to_float(self._first(data, "lat", "latitude"))
        name = self._first(data, "shipname", "name")
        sog = self._to_float(self._first(data, "sog", "speed_over_ground"))
        cog = self._to_float(self._first(data, "cog", "course_over_ground"))
        hdg_raw = self._first(data, "true_heading", "heading")
        hdg = None if hdg_raw in (None, 511) else self._to_float(hdg_raw)

        a, b, c, d = self._extract_dims(data)
        length_m = (a + b) if (a is not None and b is not None) else None
        width_m  = (c + d) if (c is not None and d is not None) else None

        if lon is None or lat is None:
            self._update_latest_attrs_in_place(mmsi, {
                "NAME": name, "SOG": sog, "COG": cog, "HDG": hdg,
                "LAST_TS": float(ts),
                "DIM_A": a, "DIM_B": b, "DIM_C": c, "DIM_D": d,
                "LENGTH_M": length_m, "WIDTH_M": width_m
            })
            if self.tracked_mmsi == mmsi:
                self._update_labels(mmsi)
                self.iface.mapCanvas().refresh()
            return

        # Clear previous latest + tracking for this MMSI, then insert new
        self._clear_latest_for_mmsi(mmsi)
        self._insert_point(mmsi, lon, lat, ts, name, sog, cog, hdg,
                           dims=(a, b, c, d), derived=(length_m, width_m))

        if self.tracked_mmsi == mmsi:
            # ensure the newest record is marked tracked
            self._set_tracking_flag_for_mmsi(mmsi, 1)
            self._update_labels(mmsi)
            self._center_on_vessel(mmsi, rezoom=False)

        self.ais.triggerRepaint()
        self.iface.mapCanvas().refresh()

    def _update_latest_attrs_in_place(self, mmsi: int, base_updates: dict):
        if not self.ais:
            return
        pr = self.ais.dataProvider()
        req = QgsFeatureRequest().setFilterExpression(f'"MMSI" = {int(mmsi)} AND "IS_LATEST" = 1')
        feats = list(self.ais.getFeatures(req))
        if not feats:
            return
        f = feats[0]
        changes = {}
        fields = self.ais.fields()
        for key, val in (base_updates or {}).items():
            if fields.indexFromName(key) >= 0:
                idx = fields.indexFromName(key)
                changes[idx] = val if val is not None else None
        if changes:
            pr.changeAttributeValues({f.id(): changes})
            self.ais.updateFields()
            self.ais.triggerRepaint()

    def _check_tracking_timeout(self):
        if not self.tracked_mmsi or not self.ais:
            return
        req = QgsFeatureRequest().setFilterExpression(f'"MMSI" = {int(self.tracked_mmsi)} AND "IS_LATEST" = 1')
        feats = list(self.ais.getFeatures(req))
        if not feats:
            return
        f = feats[0]
        try:
            last_ts = float(f["LAST_TS"])
        except Exception:
            return
        if (time.time() - (last_ts/1000.0 if last_ts >= 1e12 else last_ts)) > 600:
            # Deactivate tracking and clear flags
            if hasattr(self.dock, "TargetVessel_Checkbox"):
                try:
                    self.dock.TargetVessel_Checkbox.setChecked(False)
                except Exception:
                    pass
            self.tracked_mmsi = None
            self._clear_labels()
            self._clear_all_tracking_flags()

    def _clear_latest_for_mmsi(self, mmsi: int):
        if not self.ais:
            return
        pr = self.ais.dataProvider()
        idx_latest = self.ais.fields().indexFromName("IS_LATEST")
        idx_track  = self.ais.fields().indexFromName("IS_TRACKING")
        if idx_latest < 0:
            return
        req = QgsFeatureRequest().setFilterExpression(
            f'"MMSI" = {int(mmsi)} AND "IS_LATEST" = 1'
        ).setFlags(QgsFeatureRequest.NoGeometry)
        changes = {}
        for f in self.ais.getFeatures(req):
            ch = {idx_latest: 0}
            if idx_track >= 0:
                ch[idx_track] = 0  # also clear tracking on the previously-latest
            changes[f.id()] = ch
        if changes:
            pr.changeAttributeValues(changes)

    def _insert_point(self, mmsi, lon, lat, ts, name, sog, cog, hdg, dims=None, derived=None):
        lyr = self.ais
        pr = lyr.dataProvider()

        f = QgsFeature(lyr.fields())
        f.setGeometry(QgsGeometry.fromPointXY(QgsPointXY(lon, lat)))
        f.setAttribute("MMSI", int(mmsi))
        f.setAttribute("NAME", name if name else None)
        f.setAttribute("SOG", sog if sog is not None else None)
        f.setAttribute("COG", cog if cog is not None else None)
        f.setAttribute("HDG", hdg if hdg is not None else None)
        f.setAttribute("LAST_TS", float(ts))
        f.setAttribute("IS_LATEST", 1)
        f.setAttribute("IS_TRACKING", 1 if (self.tracked_mmsi == mmsi) else 0)

        if dims:
            a, b, c, d = dims
            if a is not None: f.setAttribute("DIM_A", a)
            if b is not None: f.setAttribute("DIM_B", b)
            if c is not None: f.setAttribute("DIM_C", c)
            if d is not None: f.setAttribute("DIM_D", d)
        if derived:
            length_m, width_m = derived
            if length_m is not None: f.setAttribute("LENGTH_M", length_m)
            if width_m  is not None: f.setAttribute("WIDTH_M",  width_m)

        pr.addFeatures([f])

    # ---------- AIS Target Selection & Labels ----------
    def on_selection_changed(self, *args):
        if not self.ais:
            return
        sel = list(self.ais.selectedFeatures())
        if not sel:
            return
        feat = sel[0]
        mmsi = feat["MMSI"]
        if mmsi is None:
            return
        self.tracked_mmsi = int(mmsi)
        self._update_labels(self.tracked_mmsi)

    def _update_labels(self, mmsi: int):
        req = QgsFeatureRequest().setFilterExpression(f'"MMSI" = {int(mmsi)} AND "IS_LATEST" = 1')
        feats = list(self.ais.getFeatures(req))
        if not feats:
            self._clear_labels()
            return
        f = feats[0]

        vessel_name = f["NAME"] if f["NAME"] not in (None, "", "NULL") else "-"
        self._set_label_text("lblVessel", vessel_name)

        self._set_label_text("lblMMSI", str(f["MMSI"]) if f["MMSI"] not in (None, "", "NULL") else "-")
        self._set_label_text("lblSOG", f"{f['SOG']} kts" if f["SOG"] not in (None, "", "NULL") else "-")
        self._set_label_text("lblCOG", f"{f['COG']}°" if f["COG"] not in (None, "", "NULL") else "-")
        self._set_label_text("lblHDG", f"{f['HDG']}°" if f["HDG"] not in (None, "", "NULL") else "-")

        # Minutes label (handles µs/ms/s)
        mins_txt = ""
        try:
            last_ts = f["LAST_TS"]
            if last_ts is None or float(last_ts) <= 0:
                mins = 0
            else:
                lt = float(last_ts)
                if lt >= 1e14:
                    last_sec = lt / 1e6
                elif lt >= 1e11:
                    last_sec = lt / 1000
                else:
                    last_sec = lt
                mins = round(max(0.0, time.time() - last_sec) / 60)
            mins_txt = f"{mins} minute(s)"
        except Exception:
            pass
        # Support both names just in case
        self._set_label_text("lblMSGRVD", mins_txt)
        self._set_label_text("lblMSGRCVD", mins_txt)

    def _clear_labels(self):
        for l in ["lblVessel", "lblMMSI", "lblSOG", "lblCOG", "lblHDG", "lblMSGRVD", "lblMSGRCVD"]:
            self._set_label_text(l, "")

    def _set_label_text(self, name, text):
        if hasattr(self.dock, name):
            try:
                getattr(self.dock, name).setText(str(text))
            except Exception:
                pass

    # ---------- AIS Target Tracking & centering ----------
    def toggle_tracking(self, state):
        if not state:
            # OFF → clear all IS_TRACKING
            self._clear_all_tracking_flags()
            self.tracked_mmsi = None
            self._clear_labels()
            if self.ais:
                self.ais.triggerRepaint()
            return

        
        if not self.tracked_mmsi:
            self._set_status("Select a vessel first")
            try:
                self.dock.TargetVessel_Checkbox.setChecked(False)
            except Exception:
                pass
            return

        # Clear all, then set IS_TRACKING = 1 for selected MMSI's latest
        self._clear_all_tracking_flags()
        self._set_tracking_flag_for_mmsi(self.tracked_mmsi, 1)

        # Deselect and center/zoom once
        try:
            self.ais.removeSelection()
        except Exception:
            pass
        self._center_on_vessel(self.tracked_mmsi, rezoom=True)

    def _clear_all_tracking_flags(self):
        if not self.ais:
            return
        pr = self.ais.dataProvider()
        idx = self.ais.fields().indexFromName("IS_TRACKING")
        if idx < 0:
            return
        changes = {}
        for f in self.ais.getFeatures(QgsFeatureRequest().setFlags(QgsFeatureRequest.NoGeometry)):
            if f["IS_TRACKING"] != 0:
                changes[f.id()] = {idx: 0}
        if changes:
            pr.changeAttributeValues(changes)

    def _set_tracking_flag_for_mmsi(self, mmsi: int, val: int = 1):
        if not self.ais:
            return
        pr = self.ais.dataProvider()
        idx = self.ais.fields().indexFromName("IS_TRACKING")
        if idx < 0:
            return
        req = QgsFeatureRequest().setFilterExpression(f'"MMSI" = {int(mmsi)} AND "IS_LATEST" = 1')
        changes = {}
        for f in self.ais.getFeatures(req):
            changes[f.id()] = {idx: int(val)}
        if changes:
            pr.changeAttributeValues(changes)

    def _center_on_vessel(self, mmsi: int, rezoom: bool = False):
        req = QgsFeatureRequest().setFilterExpression(f'"MMSI" = {int(mmsi)} AND "IS_LATEST" = 1')
        feats = list(self.ais.getFeatures(req))
        if not feats:
            return
        geom = feats[0].geometry()
        if not geom:
            return
        pt = geom.asPoint() if not geom.isMultipart() else geom.asMultiPoint()[0]
        try:
            self.iface.mapCanvas().setCenter(QgsPointXY(pt.x(), pt.y()))
        except Exception:
            self.iface.mapCanvas().setExtent(geom.boundingBox())
        if rezoom:
            try:
                self.iface.mapCanvas().zoomScale(10000)
            except Exception:
                pass
        self.iface.mapCanvas().refresh()

    # ---------- AIS Message - decode the vessel dimensions (if available) ----------
    def _extract_dims(self, data: dict):
        if not data:
            return None, None, None, None
        def g(*keys):
            for k in keys:
                if k in data and data[k] is not None:
                    try:
                        return float(data[k])
                    except Exception:
                        pass
            return None
        a = g("dim_a", "to_bow", "dim_to_bow", "dimension_to_bow")
        b = g("dim_b", "to_stern", "dim_to_stern", "dimension_to_stern")
        c = g("dim_c", "to_port", "dim_to_port", "dimension_to_port")
        d = g("dim_d", "to_starboard", "dim_to_starboard", "dimension_to_starboard")
        return a, b, c, d

    # ---------- Utility ----------
    def _get_layer_by_name(self, name):
        lst = QgsProject.instance().mapLayersByName(name)
        return lst[0] if lst else None

    def _to_int(self, v):
        try:
            return None if v in (None, "") else int(v)
        except Exception:
            return None

    def _to_float(self, v):
        try:
            return None if v in (None, "") else float(v)
        except Exception:
            return None

    def _first(self, d: dict, *keys):
        for k in keys:
            if k in d and d[k] is not None:
                return d[k]
        return None

    def _load_plugin_icon(self):
        for p in (":/plugins/QAIS/icon.png",
                  ":/plugins/QAIS/icon.svg",
                  os.path.join(os.path.dirname(__file__), "resources", "icon.png"),
                  os.path.join(os.path.dirname(__file__), "resources", "icon.svg")):
            try:
                ic = QIcon(p)
                if not ic.isNull():
                    return ic
            except Exception:
                pass
        return QIcon()
