# -*- coding: utf-8 -*-
# |1231

import pythoncom
from win32com.client import Dispatch, gencache

import LDefin2D
import MiscellaneousHelpers as MH

#  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule(
    "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0
).constants
kompas6_constants_3d = gencache.EnsureModule(
    "{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0
).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule(
    "{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0
)
kompas_object = kompas6_api5_module.KompasObject(
    Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(
        kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch
    )
)
MH.iKompasObject = kompas_object

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule(
    "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0
)
application = kompas_api7_module.IApplication(
    Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(
        kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch
    )
)
MH.iApplication = application


Documents = application.Documents
#  Создаем новый документ
kompas_document = Documents.AddWithDefaultSettings(
    kompas6_constants.ksDocumentPart, True
)

kompas_document_3d = kompas_api7_module.IKompasDocument3D(kompas_document)
iDocument3D = kompas_object.ActiveDocument3D()

iPart7 = kompas_document_3d.TopPart
iPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)

iSketch = iPart.NewEntity(kompas6_constants_3d.o3d_sketch)
iDefinition = iSketch.GetDefinition()
iPlane = iPart.GetDefaultEntity(kompas6_constants_3d.o3d_planeXOY)
iDefinition.SetPlane(iPlane)
iSketch.Create()
iDocument2D = iDefinition.BeginEdit()
kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
iDocument2D = kompas_object.ActiveDocument2D()

obj = iDocument2D.ksLineSeg(
    -17.991669807357, -7.937499999708, -7.991669807357, -7.937499999708, 1
)
obj = iDocument2D.ksLineSeg(
    -7.991669807357, -7.937499999708, -7.991669807357, 9.062500000292, 1
)
obj = iDocument2D.ksLineSeg(
    -7.991669807357, 9.062500000292, -17.991669807357, 9.062500000292, 1
)
obj = iDocument2D.ksLineSeg(
    -17.991669807357, 9.062500000292, -17.991669807357, -7.937499999708, 1
)
obj = iDocument2D.ksLineSeg(
    -7.991669807357, -7.937499999708, -17.991669807357, 9.062500000292, 2
)
obj = iDocument2D.ksLineSeg(
    -17.991669807357, -7.937499999708, -7.991669807357, 9.062500000292, 2
)
obj = iDocument2D.ksPoint(0, 0, 0)
iDefinition.EndEdit()
iDefinition.angle = 180
iSketch.Update()
iPart7 = kompas_document_3d.TopPart
iPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)

obj = iPart.NewEntity(kompas6_constants_3d.o3d_bossExtrusion)
iDefinition = obj.GetDefinition()
iCollection = iPart.EntityCollection(kompas6_constants_3d.o3d_edge)
iCollection.SelectByPoint(12.991669807357, -9.062500000292, 0)
iEdge = iCollection.Last()
iEdgeDefinition = iEdge.GetDefinition()
iSketch = iEdgeDefinition.GetOwnerEntity()
iDefinition.SetSketch(iSketch)
iExtrusionParam = iDefinition.ExtrusionParam()
iExtrusionParam.direction = kompas6_constants_3d.dtNormal
iExtrusionParam.depthNormal = 10
iExtrusionParam.depthReverse = 0
iExtrusionParam.draftOutwardNormal = False
iExtrusionParam.draftOutwardReverse = False
iExtrusionParam.draftValueNormal = 0
iExtrusionParam.draftValueReverse = 0
iExtrusionParam.typeNormal = kompas6_constants_3d.etBlind
iExtrusionParam.typeReverse = kompas6_constants_3d.etBlind
iThinParam = iDefinition.ThinParam()
iThinParam.thin = False
obj.name = "Элемент выдавливания:1"
iColorParam = obj.ColorParam()
iColorParam.ambient = 0.5
iColorParam.color = 9474192
iColorParam.diffuse = 0.6
iColorParam.emission = 0.5
iColorParam.shininess = 0.8
iColorParam.specularity = 0.8
iColorParam.transparency = 1
obj.Create()
iPart7 = kompas_document_3d.TopPart
iPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)

iSketch = iPart.NewEntity(kompas6_constants_3d.o3d_sketch)
iDefinition = iSketch.GetDefinition()
iCollection = iPart.EntityCollection(kompas6_constants_3d.o3d_face)
iCollection.SelectByPoint(12.991669807357, -0.562500000292, 10)
iPlane = iCollection.First()
iDefinition.SetPlane(iPlane)
iSketch.Create()
iDocument2D = iDefinition.BeginEdit()
kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
iDocument2D = kompas_object.ActiveDocument2D()

obj = iDocument2D.ksLine(-17.991669807357, 9.062500000292, 300.46554491946)
obj = iDocument2D.ksLine(-17.991669807357, -7.937499999708, 59.53445508054)
obj = iDocument2D.ksCircle(-12.991669807357, 0.562500000292, 3, 1)
iDefinition.EndEdit()
iDefinition.angle = 180
iSketch.Update()
iPart7 = kompas_document_3d.TopPart
iPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)

obj = iPart.NewEntity(kompas6_constants_3d.o3d_cutExtrusion)
iDefinition = obj.GetDefinition()
iCollection = iPart.EntityCollection(kompas6_constants_3d.o3d_edge)
iCollection.SelectByPoint(15.991669807357, -0.562500000292, 10)
iEdge = iCollection.Last()
iEdgeDefinition = iEdge.GetDefinition()
iSketch = iEdgeDefinition.GetOwnerEntity()
iDefinition.SetSketch(iSketch)
iDefinition.cut = True
iExtrusionParam = iDefinition.ExtrusionParam()
iExtrusionParam.direction = kompas6_constants_3d.dtNormal
iExtrusionParam.depthNormal = 10
iExtrusionParam.depthReverse = 0
iExtrusionParam.draftOutwardNormal = False
iExtrusionParam.draftOutwardReverse = False
iExtrusionParam.draftValueNormal = 0
iExtrusionParam.draftValueReverse = 0
iExtrusionParam.typeNormal = kompas6_constants_3d.etThroughAll
iExtrusionParam.typeReverse = kompas6_constants_3d.etBlind
iThinParam = iDefinition.ThinParam()
iThinParam.thin = False
obj.name = "Элемент выдавливания:2"
iColorParam = obj.ColorParam()
iColorParam.ambient = 0.5
iColorParam.color = 9474192
iColorParam.diffuse = 0.6
iColorParam.emission = 0.5
iColorParam.shininess = 0.8
iColorParam.specularity = 0.8
iColorParam.transparency = 1
obj.Create()
iPart7 = kompas_document_3d.TopPart
iPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)

obj = iPart.NewEntity(kompas6_constants_3d.o3d_fillet)
iDefinition = obj.GetDefinition()
iDefinition.radius = 1
iDefinition.tangent = True
iArray = iDefinition.array()
iCollection = iPart.EntityCollection(kompas6_constants_3d.o3d_edge)
iCollection.SelectByPoint(15.991669807357, -0.562500000292, 10)
iEdge = iCollection.Last()
iArray.Add(iEdge)
obj.name = "Скругление:1"
iColorParam = obj.ColorParam()
iColorParam.ambient = 0.5
iColorParam.color = 9474192
iColorParam.diffuse = 0.6
iColorParam.emission = 0.5
iColorParam.shininess = 0.8
iColorParam.specularity = 0.8
iColorParam.transparency = 1
obj.Create()
# kompas_document.SaveAs(r"C:\Users\artem\Desktop\Деталь.m3d")
