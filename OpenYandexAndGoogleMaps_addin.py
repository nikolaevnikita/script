import arcpy, pythonaddins, webbrowser, threading

class OpenGoogleMaps(object):
    """Implementation for OpenYandexAndGoogleMaps_addin.tool_1 (Tool)"""
    def __init__(self):
        self.enabled = True
        self.shape = "NONE" # Can set to "Line", "Circle" or "Rectangle" for interactive shape drawing and to activate the onLine/Polygon/Circle event sinks.
    def onMouseDown(self, x, y, button, shift):
        pass
    def onMouseDownMap(self, x, y, button, shift):
        pass
    def onMouseUp(self, x, y, button, shift):
        pass
    def onMouseUpMap(self, x, y, button, shift):
        mxd=arcpy.mapping.MapDocument("current")
        df = arcpy.mapping.ListDataFrames(mxd)[0]
        if df.scale < 925: z=19
        elif df.scale < 1875: z=18
        elif df.scale < 3750: z=17
        elif df.scale < 7500: z=16
        elif df.scale < 17500: z=15
        elif df.scale < 37500: z=14
        elif df.scale < 75000: z=13
        elif df.scale < 175000: z=12
        elif df.scale < 375000: z=11
        elif df.scale < 750000: z=10
        elif df.scale < 1500000: z=9
        elif df.scale < 3000000: z=8
        elif df.scale < 6000000: z=7
        elif df.scale < 11500000: z=6
        elif df.scale < 22500000: z=5
        elif df.scale < 45000000: z=4
        else: z=3  
        sr_df = df.spatialReference.PCSCode
        sr_sp = arcpy.SpatialReference(sr_df)
        sr_wgs = arcpy.SpatialReference(4326)
        pt=arcpy.PointGeometry(arcpy.Point(x,y),sr_sp).projectAs(sr_wgs)
        if shift == 2:
            address = 'www.google.com/maps/@%f,%f,18z' % (pt.centroid.Y, pt.centroid.X)
        else:
            address = 'www.google.com/maps/@%f,%f,%2fz' % (pt.centroid.Y, pt.centroid.X, z)
        address_sat = address + '/data=!3m1!1e3'
        if button == 2:
            threading.Thread(target=webbrowser.open, args=(address_sat,0)).start()
        else:
            threading.Thread(target=webbrowser.open, args=(address,0)).start()
    def onMouseMove(self, x, y, button, shift):
        pass
    def onMouseMoveMap(self, x, y, button, shift):
        pass
    def onDblClick(self):
        pass
    def onKeyDown(self, keycode, shift):
        pass
    def onKeyUp(self, keycode, shift):
        print keycode
        pass
    def deactivate(self):
        pass
    def onCircle(self, circle_geometry):
        pass
    def onLine(self, line_geometry):
        pass
    def onRectangle(self, rectangle_geometry):
        pass

class OpenYandexMaps(object):
    """Implementation for OpenYandexAndGoogleMaps_addin.tool (Tool)"""
    def __init__(self):
        self.enabled = True
        self.shape = "NONE" # Can set to "Line", "Circle" or "Rectangle" for interactive shape drawing and to activate the onLine/Polygon/Circle event sinks.
    def onMouseDown(self, x, y, button, shift):
        pass
    def onMouseDownMap(self, x, y, button, shift):
        pass
    def onMouseUp(self, x, y, button, shift):
        pass
    def onMouseUpMap(self, x, y, button, shift):
        mxd=arcpy.mapping.MapDocument("current")
        df = arcpy.mapping.ListDataFrames(mxd)[0]
        if df.scale < 925: z=19
        elif df.scale < 1875: z=18
        elif df.scale < 3750: z=17
        elif df.scale < 7500: z=16
        elif df.scale < 17500: z=15
        elif df.scale < 37500: z=14
        elif df.scale < 75000: z=13
        elif df.scale < 175000: z=12
        elif df.scale < 375000: z=11
        elif df.scale < 750000: z=10
        elif df.scale < 1500000: z=9
        elif df.scale < 3000000: z=8
        elif df.scale < 6000000: z=7
        elif df.scale < 11500000: z=6
        elif df.scale < 22500000: z=5
        elif df.scale < 45000000: z=4
        else: z=3   
        sr_df = df.spatialReference.PCSCode
        sr_sp = arcpy.SpatialReference(sr_df)
        sr_wgs = arcpy.SpatialReference(4326)
        pt=arcpy.PointGeometry(arcpy.Point(x,y),sr_sp).projectAs(sr_wgs)
        if shift == 2:
            address = 'www.yandex.ru/maps/?ll='+str(pt.centroid.X)+'%2C'+str(pt.centroid.Y)+'&z=18'
        else:
            address = 'www.yandex.ru/maps/?ll='+str(pt.centroid.X)+'%2C'+str(pt.centroid.Y)+'&z='+str(z)
        address_sat = address + '&l=sat%2Cskl'
        if button == 2:
            threading.Thread(target=webbrowser.open, args=(address_sat,0)).start()
        else:
            threading.Thread(target=webbrowser.open, args=(address,0)).start()
    def onMouseMove(self, x, y, button, shift):
        pass
    def onMouseMoveMap(self, x, y, button, shift):
        pass
    def onDblClick(self):
        pass
    def onKeyDown(self, keycode, shift):
        pass
    def onKeyUp(self, keycode, shift):
        pass
    def deactivate(self):
        pass
    def onCircle(self, circle_geometry):
        pass
    def onLine(self, line_geometry):
        pass
    def onRectangle(self, rectangle_geometry):
        pass
