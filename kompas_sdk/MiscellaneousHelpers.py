# -*- coding: cp1251 -*-
# Проект: КОМПАС-Макро
# Модуль: MiscellaneousHelpers.py (Различные функции - помошники)
# Дата создания: 07.11.2008
# Дата последнего изменения: 13.11.2008
# Автор: Архипов Александр (АСКОН-Курган)
#
import pythoncom
from win32com.client import Dispatch, gencache

kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants

# Точность сравнения равенства вещественных чисел
DoublePrecition = 1.0e-11

# Ограничение по количеству точек сравнения для ломаных
MaxNumberOfPointsToCompareArrays = 10

# Проверка на равенство 2-х вещественных чисел
def DoubleEqualTo(double1, double2):
    return abs(double1 - double2) <= 2 * DoublePrecition

# Ищет ссылку отрезка
def GetLineSegReference (X1, Y1, X2, Y2, Style):
    iDocument2D = iKompasObject.ActiveDocument2D()

    iIterator = iKompasObject.GetIterator()
    iIterator.ksCreateIterator(kompas6_constants.ksDrLineSeg, 0)
    refLine = iIterator.ksMoveIterator('F')
    iksLineSegParam = iKompasObject.GetParamStruct(kompas6_constants.ko_LineSegParam)

    while refLine != 0:
        iDocument2D.ksGetObjParam(refLine, iksLineSegParam, kompas6_constants.ksAllParam)
        if DoubleEqualTo(iksLineSegParam.x1, X1) and \
           DoubleEqualTo(iksLineSegParam.y1, Y1) and \
           DoubleEqualTo(iksLineSegParam.x2, X2) and \
           DoubleEqualTo(iksLineSegParam.y2, Y2) and \
           iksLineSegParam.style == Style : \
           break
        refLine = iIterator.ksMoveIterator('N')

    iIterator.ksDeleteIterator()
    return refLine


# Ищет ссылку линии
def GetLineReference (X, Y, Angle):
    iDocument2D = iKompasObject.ActiveDocument2D()

    iIterator = iKompasObject.GetIterator()
    iIterator.ksCreateIterator(kompas6_constants.ksDrLine, 0)
    refLine = iIterator.ksMoveIterator('F')
    iksLineParam = iKompasObject.GetParamStruct(kompas6_constants.ko_LineParam)

    while refLine != 0:
        iDocument2D.ksGetObjParam(refLine, iksLineParam, kompas6_constants.ksAllParam)
        if DoubleEqualTo(iksLineParam.x, X) and \
           DoubleEqualTo(iksLineParam.y, Y) and \
           DoubleEqualTo(iksLineParam.angle, Angle): \
           break
        refLine = iIterator.ksMoveIterator('N')

    iIterator.ksDeleteIterator()
    return refLine

# Ищет ссылку окружности
def GetCircleReference (X, Y, Radius, Style):
    iDocument2D = iKompasObject.ActiveDocument2D()

    iIterator = iKompasObject.GetIterator()
    iIterator.ksCreateIterator(kompas6_constants.ksDrCircle, 0)
    refLine = iIterator.ksMoveIterator('F')
    iksCircleParam = iKompasObject.GetParamStruct(kompas6_constants.ko_CircleParam)

    while refLine != 0:
        iDocument2D.ksGetObjParam(refLine, iksCircleParam, kompas6_constants.ksAllParam)
        if DoubleEqualTo(iksCircleParam.xc, X) and \
           DoubleEqualTo(iksCircleParam.yc, Y) and \
           DoubleEqualTo(iksCircleParam.rad, Radius) and \
           iksCircleParam.style == Style: \
           break
        refLine = iIterator.ksMoveIterator('N')

    iIterator.ksDeleteIterator()
    return refLine

# Ищет ссылку прямоугольника
def GetRectangleReference (X, Y, Angle, Height, Width, Style):
    iDocument2D = iKompasObject.ActiveDocument2D()

    iIterator = iKompasObject.GetIterator()
    iIterator.ksCreateIterator(kompas6_constants.ksDrRectangle, 0)
    refLine = iIterator.ksMoveIterator('F')
    iksRectangleParam = iKompasObject.GetParamStruct(kompas6_constants.ko_RectangleParam)

    while refLine != 0:
        iDocument2D.ksGetObjParam(refLine, iksRectangleParam, kompas6_constants.ksAllParam)
        if DoubleEqualTo(iksRectangleParam.x, X) and \
           DoubleEqualTo(iksRectangleParam.y, Y) and \
           DoubleEqualTo(iksRectangleParam.ang, Angle) and \
           DoubleEqualTo(iksRectangleParam.height, Height) and \
           DoubleEqualTo(iksRectangleParam.width, Width) and \
           iksRectangleParam.style == Style: \
           break

        refLine = iIterator.ksMoveIterator('N')

    iIterator.ksDeleteIterator()
    return refLine

# Ищет ссылку эллипса
def GetEllipseReference (X, Y, A, B, Angle, Style):
    iDocument2D = iKompasObject.ActiveDocument2D()

    iIterator = iKompasObject.GetIterator()
    iIterator.ksCreateIterator(kompas6_constants.ksDrEllipse, 0)
    refLine = iIterator.ksMoveIterator('F')
    iksEllipseParam = iKompasObject.GetParamStruct(kompas6_constants.ko_EllipseParam)

    while refLine != 0:
        iDocument2D.ksGetObjParam(refLine, iksEllipseParam, kompas6_constants.ksAllParam)
        if DoubleEqualTo(iksEllipseParam.xc, X) and \
           DoubleEqualTo(iksEllipseParam.yc, Y) and \
           DoubleEqualTo(iksEllipseParam.A, A) and \
           DoubleEqualTo(iksEllipseParam.B, B) and \
           DoubleEqualTo(iksEllipseParam.angle, Angle) and \
           iksEllipseParam.style == Style: \
           break

        refLine = iIterator.ksMoveIterator('N')

    iIterator.ksDeleteIterator()
    return refLine

# Ищет ссылку дуги
def GetArcReference (X, Y, Radius, Angle1, Angle2, Direction, Style):
    iDocument2D = iKompasObject.ActiveDocument2D()

    iIterator = iKompasObject.GetIterator()
    iIterator.ksCreateIterator(kompas6_constants.ksDrArc, 0)
    refLine = iIterator.ksMoveIterator('F')
    iksArcByAngleParam = iKompasObject.GetParamStruct(kompas6_constants.ko_ArcByAngleParam)

    while refLine != 0:
        iDocument2D.ksGetObjParam(refLine, iksArcByAngleParam, kompas6_constants.ksAllParam)
        if DoubleEqualTo(iksArcByAngleParam.xc, X) and \
           DoubleEqualTo(iksArcByAngleParam.yc, Y) and \
           DoubleEqualTo(iksArcByAngleParam.rad, Radius) and \
           DoubleEqualTo(iksArcByAngleParam.ang1, Angle1) and \
           DoubleEqualTo(iksArcByAngleParam.ang2, Angle2) and \
           iksArcByAngleParam.dir == Direction and \
           iksArcByAngleParam.style == Style: \
           break

        refLine = iIterator.ksMoveIterator('N')

    iIterator.ksDeleteIterator()
    return refLine

# Сравнение массива точек типа ksDynamicArray с данным кортежем точек
# число сравниваемых точек не превышает MaximumNumberOfPoints
def DynamicArrayHasPoints(iksDynamicArray, ko_PointParam, Points, MaximumNumberOfPoints):
    if iksDynamicArray.ksGetArrayCount() != len(Points) :
        return False
    for i in range(min(len(Points), MaximumNumberOfPoints)):
        iksPointParam = iKompasObject.GetParamStruct(ko_PointParam)
        iksDynamicArray.ksGetArrayItem(i, iksPointParam)
        if not DoubleEqualTo(iksPointParam.x, Points[i][0]) or not DoubleEqualTo(iksPointParam.y, Points[i][1]) :
            return False
    return True

# Ищет ссылку ломаной
def GetPolylineReference (Points):
    iDocument2D = iKompasObject.ActiveDocument2D()

    iIterator = iKompasObject.GetIterator()
    iIterator.ksCreateIterator(kompas6_constants.ksDrPolyline, 0)
    refLine = iIterator.ksMoveIterator('F')
    iksPolylineParam = iKompasObject.GetParamStruct(kompas6_constants.ko_PolylineParam)

    while refLine != 0:
        iDocument2D.ksGetObjParam(refLine, iksPolylineParam, kompas6_constants.ksAllParam)
        iksDynamicArray = iksPolylineParam.GetpMathPoint()
        if DynamicArrayHasPoints(iksDynamicArray, kompas6_constants.ko_MathPointParam, Points, MaxNumberOfPointsToCompareArrays) : break

        refLine = iIterator.ksMoveIterator('N')

    iIterator.ksDeleteIterator()
    return refLine


# Ищет ссылку сплайна Безье
def GetBezierSplineReference (Points):
    iDocument2D = iKompasObject.ActiveDocument2D()

    iIterator = iKompasObject.GetIterator()
    iIterator.ksCreateIterator(kompas6_constants.ksDrBezier, 0)
    refLine = iIterator.ksMoveIterator('F')
    iksBezierParam = iKompasObject.GetParamStruct(kompas6_constants.ko_BezierParam)

    while refLine != 0:
        iDocument2D.ksGetObjParam(refLine, iksBezierParam, kompas6_constants.ksAllParam)
        iksDynamicArray = iksBezierParam.GetMathPointArr()
        if DynamicArrayHasPoints(iksDynamicArray, kompas6_constants.ko_MathPointParam, Points, MaxNumberOfPointsToCompareArrays) : break

        refLine = iIterator.ksMoveIterator('N')

    iIterator.ksDeleteIterator()
    return refLine


"""
# Ищет ссылку стрелки направления взгляда
def GetWPointerReference (X1, Y1, X2, Y2, XT, YT, Type, String, Style):
    iDocument2D = iKompasObject.ActiveDocument2D()

    iIterator = iKompasObject.GetIterator()
    iIterator.ksCreateIterator(kompas6_constants.ksDrWPointer, 0)
    refLine = iIterator.ksMoveIterator('F')
    iksViewPointerParam = iKompasObject.GetParamStruct(kompas6_constants.ko_ViewPointerParam)

    while refLine != 0:
        iDocument2D.ksGetObjParam(refLine, iksViewPointerParam, kompas6_constants.ksAllParam)
        if  DoubleEqualTo(iksViewPointerParam.x1, X1) and \
            DoubleEqualTo(iksViewPointerParam.y1, Y1) and \
            DoubleEqualTo(iksViewPointerParam.x2, X2) and \
            DoubleEqualTo(iksViewPointerParam.y2, Y2) and \
            DoubleEqualTo(iksViewPointerParam.xt, XT) and \
            DoubleEqualTo(iksViewPointerParam.yt, YT) and \
            iksViewPointerParam.type == Type and \
            iksViewPointerParam.str == String and \
            iksViewPointerParam.style == Style :
            break

        refLine = iIterator.ksMoveIterator('N')

    iIterator.ksDeleteIterator()
    return refLine
"""

# Ищет ссылку NURBS сплайна
def GetNurbsSplineReference (Points):
    iDocument2D = iKompasObject.ActiveDocument2D()

    iIterator = iKompasObject.GetIterator()
    iIterator.ksCreateIterator(kompas6_constants.ksDrNurbs, 0)
    refLine = iIterator.ksMoveIterator('F')
    iksNurbsParam = iKompasObject.GetParamStruct(kompas6_constants.ko_NurbsParam)

    while refLine != 0:
        iDocument2D.ksGetObjParam(refLine, iksNurbsParam, kompas6_constants.ksAllParam)
        iksDynamicArray = iksNurbsParam.GetPPoint()
        if DynamicArrayHasPoints(iksDynamicArray, kompas6_constants.ko_NurbsPointParam, Points, MaxNumberOfPointsToCompareArrays) : break

        refLine = iIterator.ksMoveIterator('N')

    iIterator.ksDeleteIterator()
    return refLine


# Ищет эквидистанту
"""
def GetEquidistantReference (X1, Y1, X2, Y2, XT, YT, Type, String, Style):
    iDocument2D = iKompasObject.ActiveDocument2D()

    iIterator = iKompasObject.GetIterator()
    iIterator.ksCreateIterator(kompas6_constants.ksDrWPointer, 0)
    refLine = iIterator.ksMoveIterator('F')
    iksViewPointerParam = iKompasObject.GetParamStruct(kompas6_constants.ko_ViewPointerParam)

    while refLine != 0:
        iDocument2D.ksGetObjParam(refLine, iksViewPointerParam, kompas6_constants.ksAllParam)
        if  DoubleEqualTo(iksViewPointerParam.x1, X1) and \
            DoubleEqualTo(iksViewPointerParam.y1, Y1) and \
            DoubleEqualTo(iksViewPointerParam.x2, X2) and \
            DoubleEqualTo(iksViewPointerParam.y2, Y2) and \
            DoubleEqualTo(iksViewPointerParam.xt, XT) and \
            DoubleEqualTo(iksViewPointerParam.yt, YT) and \
            iksViewPointerParam.type == Type and \
            iksViewPointerParam.str == String and \
            iksViewPointerParam.style == Style :
            break

        refLine = iIterator.ksMoveIterator('N')

    iIterator.ksDeleteIterator()
    return refLine
"""

# Ищет ссылку осевой линии
def GetAxisLineReference (X1, Y1, X2, Y2):
    iDocument2D = iKompasObject.ActiveDocument2D()

    iIterator = iKompasObject.GetIterator()
    iIterator.ksCreateIterator(kompas6_constants.ksDrAxisLine, 0)
    refLine = iIterator.ksMoveIterator('F')
    iksAxisLineParam = iKompasObject.GetParamStruct(kompas6_constants.ko_AxisLineParam)

    while refLine != 0:
        iDocument2D.ksGetObjParam(refLine, iksAxisLineParam, kompas6_constants.ksAllParam)
        if DoubleEqualTo(iksAxisLineParam.GetBegPoint().x, X1) and \
           DoubleEqualTo(iksAxisLineParam.GetBegPoint().y, Y1) and \
           DoubleEqualTo(iksAxisLineParam.GetEndPoint().x, X2) and \
           DoubleEqualTo(iksAxisLineParam.GetEndPoint().y, Y2) :
           break

        refLine = iIterator.ksMoveIterator('N')

    iIterator.ksDeleteIterator()
    return refLine


# Ищет интерфейс прямой координационной оси
def GetStraightAxis (X1, Y1, X2, Y2, Length, Angle):
    iDocument = iApplication.ActiveDocument
    iKompasDocument2D = KAPI7.IKompasDocument2D(iDocument)

    iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager
    iViews = iViewsAndLayersManager.Views
    iView = iViews.ActiveView
    iBuildingContainer = iView._oleobj_.QueryInterface(KAPI7.NamesToIIDMap['IBuildingContainer'], pythoncom.IID_IDispatch)
    iBuildingContainer = KAPI7.IBuildingContainer(iBuildingContainer)
    iBuildingAxes = iBuildingContainer.BuildingAxes
    iStraightAxis = 0
    for i in range(iBuildingAxes.Count) :
        iStraightAxis = iBuildingAxes.BuildingAxis(i)
        if iStraightAxis.Type == kompas6_constants.ksObjectStraightAxis :
            iStraightAxis = KAPI7.IStraightAxis(iStraightAxis)
            if DoubleEqualTo(iStraightAxis.X1, X1) and\
               DoubleEqualTo(iStraightAxis.Y1, Y1) and\
               DoubleEqualTo(iStraightAxis.X2, X2) and\
               DoubleEqualTo(iStraightAxis.Y2, Y2) and\
               DoubleEqualTo(iStraightAxis.Length, Length) and\
               DoubleEqualTo(iStraightAxis.Angle, Angle):
                break
    return iStraightAxis


# Ищет интерфейс круговой координационной оси
def GetCircleAxis (X, Y, Radius):
    iDocument = iApplication.ActiveDocument
    iKompasDocument2D = KAPI7.IKompasDocument2D(iDocument)

    iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager
    iViews = iViewsAndLayersManager.Views
    iView = iViews.ActiveView
    iBuildingContainer = iView._oleobj_.QueryInterface(KAPI7.NamesToIIDMap['IBuildingContainer'], pythoncom.IID_IDispatch)
    iBuildingContainer = KAPI7.IBuildingContainer(iBuildingContainer)
    iBuildingAxes = iBuildingContainer.BuildingAxes
    iCircleAxis = 0
    for i in range(iBuildingAxes.Count) :
        iCircleAxis = iBuildingAxes.BuildingAxis(i)
        if iCircleAxis.Type == kompas6_constants.ksObjectCircleAxis :
            iCircleAxis = KAPI7.ICircleAxis(iCircleAxis)
            if DoubleEqualTo(iCircleAxis.Xc, X) and\
               DoubleEqualTo(iCircleAxis.Yc, Y) and\
               DoubleEqualTo(iCircleAxis.Radius, Radius):
                break
    return iCircleAxis


# Ищет интерфейс дуговой координационной оси
def GetArcAxis (X1, Y1, X2, Y2, X3, Y3, Xc, Yc, Radius, Angle1, Angle2, Direction):
    iDocument = iApplication.ActiveDocument
    iKompasDocument2D = KAPI7.IKompasDocument2D(iDocument)

    iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager
    iViews = iViewsAndLayersManager.Views
    iView = iViews.ActiveView
    iBuildingContainer = iView._oleobj_.QueryInterface(KAPI7.NamesToIIDMap['IBuildingContainer'], pythoncom.IID_IDispatch)
    iBuildingContainer = KAPI7.IBuildingContainer(iBuildingContainer)
    iBuildingAxes = iBuildingContainer.BuildingAxes
    iArcAxis = 0
    for i in range(iBuildingAxes.Count) :
        iArcAxis = iBuildingAxes.BuildingAxis(i)
        if iArcAxis.Type == kompas6_constants.ksObjectArcAxis :
            iArcAxis= KAPI7.IArcAxis(iArcAxis)
            if DoubleEqualTo(iArcAxis.X1, X1) and\
               DoubleEqualTo(iArcAxis.Y1, Y1) and\
               DoubleEqualTo(iArcAxis.X2, X2) and\
               DoubleEqualTo(iArcAxis.Y2, Y2) and\
               DoubleEqualTo(iArcAxis.X3, X3) and\
               DoubleEqualTo(iArcAxis.Y3, Y3) and\
               DoubleEqualTo(iArcAxis.Xc, Xc) and\
               DoubleEqualTo(iArcAxis.Yc, Yc) and\
               DoubleEqualTo(iArcAxis.Radius, Radius) and\
               DoubleEqualTo(iArcAxis.Angle1, Angle1) and\
               DoubleEqualTo(iArcAxis.Angle2, Angle2) and\
               iArcAxis.Direction == Direction :
                break
    return iArcAxis
