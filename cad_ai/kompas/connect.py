import pythoncom
from win32com.client import Dispatch, gencache

import MiscellaneousHelpers as MH

GUID_KS_CONST = "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}"
GUID_KS_CONST_3D = "{2CAF168C-7961-4B90-9DA2-701419BEEFE3}"
GUID_API5 = "{0422828C-F174-495E-AC5D-D31014DBBE87}"
GUID_API7 = "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}"


def connect_kompas():
    """
    Returns:
    (ks_const, ks_const_3d, api5, api7, kompas_object, application)
    """
    ks_const = gencache.EnsureModule(GUID_KS_CONST, 0, 1, 0).constants
    ks_const_3d = gencache.EnsureModule(GUID_KS_CONST_3D, 0, 1, 0).constants

    api5 = gencache.EnsureModule(GUID_API5, 0, 1, 0)
    api7 = gencache.EnsureModule(GUID_API7, 0, 1, 0)

    kompas5_disp = Dispatch("Kompas.Application.5")
    kompas7_disp = Dispatch("Kompas.Application.7")

    kompas_object = api5.KompasObject(
        kompas5_disp._oleobj_.QueryInterface(
            api5.KompasObject.CLSID, pythoncom.IID_IDispatch
        )
    )
    application = api7.IApplication(
        kompas7_disp._oleobj_.QueryInterface(
            api7.IApplication.CLSID, pythoncom.IID_IDispatch
        )
    )

    MH.iKompasObject = kompas_object
    MH.iApplication = application

    try:
        MH.iKompasObject.Visible = True
    except Exception:
        pass

    return ks_const, ks_const_3d, api5, api7, kompas_object, application


def new_document_part(ks_const, ks_const_3d, api5, api7, kompas_object, application):
    """
    Returns:
    (kompas_document, kompas_document_3d, iDocument3D, iPart)
    """
    documents = application.Documents
    kompas_document = documents.AddWithDefaultSettings(ks_const.ksDocumentPart, True)
    kompas_document_3d = api7.IKompasDocument3D(kompas_document)

    iDocument3D = kompas_object.ActiveDocument3D()
    iPart = iDocument3D.GetPart(ks_const_3d.pTop_Part)

    return kompas_document, kompas_document_3d, iDocument3D, iPart
