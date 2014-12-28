using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Diagnostics;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using SolidWorksTools;
using SolidWorksTools.File;

using NLog;
using NLog.Config;
using NLog.Targets;


namespace GreyWind.SolidPlayGround
{
    [Guid("2e477c75-04c5-45a1-ad64-b981ac8c62ad"), ComVisible(true)]
    [SwAddin(
        Description = "A PlayGround to test new things",
        Title = "GreyWind PlayGround",
        LoadAtStartup = true
        )]
    public class SolidWorksAddin : ISwAddin
    {
        #region Local Variables

        // logger class
        private static Logger logger = null;
        private static LogForm logForm = null;

        // Command Group
        private const int SWCommandGroupID = 318008;

        // Menu ItemIDs
        private const int mainItemID1 = 5;
        private const int mainItemID2 = 6;
        private const int mainItemID3 = 7;


        private ISldWorks _swApplication = null;
        private ICommandManager _swCommandManager = null;
        private int _swAddinID = 0;

        #region Event Handler Variables
        private Hashtable _swOpenDocs = new Hashtable();
        private SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;
        #endregion

        // Public Properties
        public ISldWorks SwApp
        {
            get { return _swApplication; }
        }
        public ICommandManager CmdMgr
        {
            get { return _swCommandManager; }
        }
        public Hashtable OpenDocs
        {
            get { return _swOpenDocs; }
        }

        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SolidWorksAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation
        public SolidWorksAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            _swApplication = (ISldWorks)ThisSW;
            _swAddinID = cookie;

            //Setup callbacks
            _swApplication.SetAddinCallbackInfo(0, this, _swAddinID);

            #region Setup Logger

            // instantiate log window, otherwise, nlog will create a window of its own.
            logForm = new LogForm();
            logForm.Show();
            logForm.Hide();

            SetUpLogger();
            logger = LogManager.GetCurrentClassLogger();
            logger.Info("AddIn loaded successfully.");
            
            #endregion

            #region Setup the Command Manager

            _swCommandManager = _swApplication.GetCommandManager(cookie);
            AddCommandMgr();
            
            #endregion

            #region Setup the Event Handlers
            
            SwEventPtr = (SldWorks)_swApplication;
            _swOpenDocs = new Hashtable();
            AttachEventHandlers();
            
            #endregion

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();
            DetachEventHandlers();

            logger.Info("Disconnecting AddIn...");

            System.Runtime.InteropServices.Marshal.ReleaseComObject(_swCommandManager);
            _swCommandManager = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_swApplication);
            _swApplication = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            logger.Info("Disconnected.");
            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            Assembly thisAssembly = Assembly.GetAssembly(this.GetType());
            BitmapHandler bitmapHandler = new BitmapHandler();

            ICommandGroup cmdGroup;
            int cmdIndex0, cmdIndex1, cmdIndex2;
            string Title = "PlayGround", ToolTip = "GreyWind PlayGround";


            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};

            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = _swCommandManager.GetGroupDataFromRegistry(SWCommandGroupID, out registryIDs);

            int[] knownIDs = new int[3] { mainItemID1, mainItemID2, mainItemID3};

            if (getDataResult)
            {
                //if the IDs don't match, reset the commandGroup
                if (!CompareIDs((int[])registryIDs, knownIDs)) 
                {
                    logger.Warn("Command Manager Buttons need to be updated!");
                    ignorePrevious = true;
                }
            }


            cmdGroup = _swCommandManager.CreateCommandGroup2(SWCommandGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);
            cmdGroup.LargeIconList = bitmapHandler.CreateFileFromResourceBitmap("GreyWind.SolidPlayGround.Images.ToolbarLarge.bmp", thisAssembly);
            cmdGroup.SmallIconList = bitmapHandler.CreateFileFromResourceBitmap("GreyWind.SolidPlayGround.Images.ToolbarSmall.bmp", thisAssembly);
            cmdGroup.LargeMainIcon = bitmapHandler.CreateFileFromResourceBitmap("GreyWind.SolidPlayGround.Images.MainIconLarge.bmp", thisAssembly);
            cmdGroup.SmallMainIcon = bitmapHandler.CreateFileFromResourceBitmap("GreyWind.SolidPlayGround.Images.MainIconSmall.bmp", thisAssembly);

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            int menuOption = (int)(swCommandItemType_e.swMenuItem);

            cmdIndex0 = cmdGroup.AddCommandItem2("Command #1", -1, "Command HintString", "Command ToolTip", 1, "STDCallBack", "", mainItemID1, menuToolbarOption);
            cmdIndex1 = cmdGroup.AddCommandItem2("Show Log", -1, "Command HintString", "Command ToolTip", 1, "ShowLog", "EnableShowLog", mainItemID2, menuOption);
            cmdIndex2 = cmdGroup.AddCommandItem2("About", -1, "About PlayGround", "About PlayGround", 3, "ShowAboutBox", "", mainItemID3, menuToolbarOption);

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();
            bitmapHandler.Dispose();


            bool bResult;

            foreach (int type in docTypes)
            {
                CommandTab cmdTab;

                cmdTab = _swCommandManager.GetCommandTab(type, Title);

                // if tab exists, but we have ignored the registry info (or changed command group ID)
                // re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                if (cmdTab != null & !getDataResult | ignorePrevious)
                {
                    bool res = _swCommandManager.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (cmdTab == null)
                {
                    cmdTab = _swCommandManager.AddCommandTab(type, Title);

                    CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

                    int[] cmdIDs = new int[3];
                    int[] TextType = new int[3];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex0);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex1);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex2);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    bResult = cmdBox.AddCommands(cmdIDs, TextType);

                }

            }
            thisAssembly = null;

        }

        public void RemoveCommandMgr()
        {
            _swCommandManager.RemoveCommandGroup(SWCommandGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {

                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        #endregion

        #region UI Callbacks

        // Nothing to see here, yet.
        public void STDCallBack()
        {
            _swApplication.SendMsgToUser("Not implemented yet!");
        }

        public void ShowLog()
        {
            logForm.Show();
        }

        public int EnableShowLog()
        {
            if (logForm.Visible == true)
            {
                return 0;
            }
            return 1;
        }

        public void ShowAboutBox()
        {
            AboutBox ab = new AboutBox();
            SwAddinAttribute SWattr = null;
            Type type = typeof(SolidWorksAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }
            ab.AppTitle = SWattr.Title;
            ab.AppDescription = SWattr.Description;

            ab.ShowDialog();
        }

        #endregion

        #region Event Methods
        public bool AttachEventHandlers()
        {
            AttachSwEvents();
            //Listen for events on all currently open docs
            AttachEventsToAllDocuments();
            return true;
        }

        private bool AttachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify += new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 += new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify += new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify += new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }



        private bool DetachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify -= new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 -= new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify -= new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify -= new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

        }

        public void AttachEventsToAllDocuments()
        {
            ModelDoc2 modDoc = (ModelDoc2)_swApplication.GetFirstDocument();
            while (modDoc != null)
            {
                if (!_swOpenDocs.Contains(modDoc))
                {
                    AttachModelDocEventHandler(modDoc);
                }
                modDoc = (ModelDoc2)modDoc.GetNext();
            }
        }

        public bool AttachModelDocEventHandler(ModelDoc2 modDoc)
        {
            if (modDoc == null)
                return false;

            DocumentEventHandler docHandler = null;

            if (!_swOpenDocs.Contains(modDoc))
            {
                switch (modDoc.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            docHandler = new PartEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            docHandler = new AssemblyEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        {
                            docHandler = new DrawingEventHandler(modDoc, this);
                            break;
                        }
                    default:
                        {
                            return false; //Unsupported document type
                        }
                }
                docHandler.AttachEventHandlers();
                _swOpenDocs.Add(modDoc, docHandler);
            }
            return true;
        }

        public bool DetachModelEventHandler(ModelDoc2 modDoc)
        {
            DocumentEventHandler docHandler;
            docHandler = (DocumentEventHandler)_swOpenDocs[modDoc];
            _swOpenDocs.Remove(modDoc);
            modDoc = null;
            docHandler = null;
            return true;
        }

        public bool DetachEventHandlers()
        {
            DetachSwEvents();

            //Close events on all currently open docs
            DocumentEventHandler docHandler;
            int numKeys = _swOpenDocs.Count;
            object[] keys = new Object[numKeys];

            //Remove all document event handlers
            _swOpenDocs.Keys.CopyTo(keys, 0);
            foreach (ModelDoc2 key in keys)
            {
                docHandler = (DocumentEventHandler)_swOpenDocs[key];
                docHandler.DetachEventHandlers(); //This also removes the pair from the hash
                docHandler = null;
            }
            return true;
        }
        #endregion

        #region Event Handlers
        //Events
        public int OnDocChange()
        {
            return 0;
        }

        public int OnDocLoad(string docTitle, string docPath)
        {
            return 0;
        }

        int FileOpenPostNotify(string FileName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnFileNew(object newDoc, int docType, string templateName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnModelChange()
        {
            return 0;
        }

        #endregion

        #region Logger Utility Methods

        private void SetUpLogger()
        {
            LoggingConfiguration logConfig = new LoggingConfiguration();

            RichTextBoxTarget textBoxTarget = new RichTextBoxTarget();
            textBoxTarget.FormName = "LogForm";
            textBoxTarget.ControlName = "LogTextBox";
            textBoxTarget.AutoScroll = true;
            textBoxTarget.UseDefaultRowColoringRules = true;
            logConfig.AddTarget("textbox", textBoxTarget);
            LoggingRule ruleTextBox = new LoggingRule("*", LogLevel.Trace, textBoxTarget);
            logConfig.LoggingRules.Add(ruleTextBox);

            FileTarget fileTarget = new FileTarget();
            fileTarget.FileName = @"${specialfolder:dir=Logs:folder=LocalApplicationData}/log.txt";
            fileTarget.Layout = @"${date:format=HH\:MM\:ss} ${callsite} ${message}";
            logConfig.AddTarget("file", fileTarget);
            LoggingRule ruleFile = new LoggingRule("*", LogLevel.Trace, fileTarget);
            logConfig.LoggingRules.Add(ruleFile);

            LogManager.Configuration = logConfig;
        }

        #endregion
    }
}
