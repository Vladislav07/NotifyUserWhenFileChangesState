using EPDM.Interop.epdm;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NotifyUserWhenFileChangesState
{
    [Guid("B4B49E1D-5CA3-4627-B3D9-E6AB286C8C50"), ComVisible(true)]
    public class ChangeStateNotification : IEdmAddIn5
    {
        SldWorks app = null;
        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {

            try
            {
                poInfo.mbsAddInName = "C# Add-In_PDM Reference file ";
                poInfo.mbsCompany = "CUBY";
                poInfo.mbsDescription = "...";
                poInfo.mlAddInVersion = 1;
                poInfo.mlRequiredVersionMajor = 6;
                poInfo.mlRequiredVersionMinor = 4;

                poCmdMgr.AddHook(EdmCmdType.EdmCmd_PostAdd);
                poCmdMgr.AddHook(EdmCmdType.EdmCmd_PreAdd);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            string pathRootName = null;
            try
            {
                switch (poCmd.meCmdType)
                {                
                    case EdmCmdType.EdmCmd_PreAdd:
                        app = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                        ModelDoc2 model = app.IActiveDoc2;
                        pathRootName = model.GetPathName();

                        break;
                    case EdmCmdType.EdmCmd_PostAdd:
                          foreach (EdmCmdData AffectedFile in ppoData)
                        {
                            string fullPath = AffectedFile.mbsStrData1;
                            long iSnetworkFiles = AffectedFile.mlLongData1;
                            string ext = Path.GetExtension(fullPath);
                            if (ext == "dxf" || ext == "DXF")
                            {
                                // TODO
                            }
                        }        
                        break;
                    default:
                       // ((EdmVault5)(poCmd.mpoVault)).MsgBox(poCmd.mlParentWnd, "An unknown command type was issued.");
                        break;
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void AddCustomFileReference(string RootPathName, string[]refFiles)
        {

 
            try
            {
  
                IEdmVault7 vault2 = null;         
                IEdmVault5  vault1 = new EdmVault5();
                vault2 = (IEdmVault7)vault1;
                if (!vault1.IsLoggedIn)
                {
                    vault1.LoginAuto("My", 0);
                }

                IEdmAddCustomRefs addCustRefs = (IEdmAddCustomRefs)vault2.CreateUtility(EdmUtility.EdmUtil_AddCustomRefs);
                IEdmFile5 rootFile = null;
                IEdmFolder5 parentFolder = null;
   
                rootFile = vault2.GetFileFromPath(RootPathName, out parentFolder);
                if (!rootFile.IsLocked)
                {
                    rootFile.LockFile(parentFolder.ID, 0, (int)EdmLockFlag.EdmLock_Simple);
                }
                
                Boolean retCode = false;

                addCustRefs.AddReferencesPath(rootFile.ID, ref refFiles);
                addCustRefs.CreateTree((int)EdmCreateReferenceFlags.Ecrf_Nothing);
                addCustRefs.ShowDlg(0);
                retCode = addCustRefs.CreateReferences();
                
                // Check in the file
                rootFile.UnlockFile(0, "Custom reference added");
                int[] ppoFileIdArray = null;
                ppoFileIdArray[0] = rootFile.ID;
                //Display current custom file references
                retCode = addCustRefs.ShowEditReferencesDlg(ref ppoFileIdArray, 0);

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
