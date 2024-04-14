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

        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {

            try
            {
                poInfo.mbsAddInName = "C# Add-In";
                poInfo.mbsCompany = "Dassault Systemes";
                poInfo.mbsDescription = "Exercise demonstrating responding to a change state event.";
                poInfo.mlAddInVersion = 1;

                //Minimum SOLIDWORKS PDM Professional version
                //needed for C# Add-Ins is 6.4
                poInfo.mlRequiredVersionMajor = 6;
                poInfo.mlRequiredVersionMinor = 4;

                //Register to receive a notification when
                //a file has changed state
                poCmdMgr.AddHook(EdmCmdType.EdmCmd_PostAdd);
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
            try
            {

                switch (poCmd.meCmdType)
                {
                    //A file has changed state
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
                    //The event isn't registered
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

        public void AddCustomFileReference(string FileName)
        {

            ModelDoc2 swModel;
            Component2 swComponent;
            SldWorks swApp;
            try
            {
                swApp = GetObject
                swModel = (ModelDoc2)swApp.ActiveDoc;

                IEdmVault7 vault2 = null;
           
                IEdmVault5  vault1 = new EdmVault5();
                vault2 = (IEdmVault7)vault1;
                if (!vault1.IsLoggedIn)
                {
                    vault1.LoginAuto("My", 0);
                }

                IEdmAddCustomRefs addCustRefs = (IEdmAddCustomRefs)vault2.CreateUtility(EdmUtility.EdmUtil_AddCustomRefs);

                IEdmFile5 file = null;
                IEdmFolder5 parentFolder = null;
   
                file = vault2.GetFileFromPath(FileName, out parentFolder);
                if (!file.IsLocked)
                {
                    file.LockFile(parentFolder.ID, 0, (int)EdmLockFlag.EdmLock_Simple);
                }
 
                
                Boolean retCode = false;
                

  
                addCustRefs.AddReferencesClipboard(file.ID);
                addCustRefs.CreateTree((int)EdmCreateReferenceFlags.Ecrf_Nothing);
                addCustRefs.ShowDlg(0);
                retCode = addCustRefs.CreateReferences();
                
                // Check in the file
                file.UnlockFile(0, "Custom reference added");
               //-- int[] ppoFileIdArray = null;
               //-- ppoFileIdArray[0] = file.ID;
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
