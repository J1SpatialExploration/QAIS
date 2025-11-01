def classFactory(iface):
    from .QAIS import QAISPlugin
    return QAISPlugin(iface)
