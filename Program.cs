using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Diagnostics;
using System.Runtime.InteropServices;

// Namespace reference to unmanaged COM library
using Com.Interwoven.WorkSite.iManage;

namespace Multi_Threaded_Search
{
    [GuidAttribute("8EF40B19-7E64-4703-A2AE-809F7AF67924"), ComVisible(true)]
    class SortWorkspaceClass : IManObjectSort
    {
        public bool Less(IManObject obj1, IManObject obj2)
        {
            IManWorkspace work1 = (IManWorkspace)obj1, work2 = (IManWorkspace)obj2;

            if (String.Compare(work1.Name.ToString(), work2.Name.ToString(), true) < 0)         //  int Compare(string strA, string strB, bool ignoreCase)
                return true;
            else
                return false;
        }
    }

    class Program
    {
        const string    strServer = "win2012svr",
                        strUsername = "wsadmin",
                        strPassword = "mhdocs_",
                        strDatabase = "Cheapside";
        private IManDMS _objDMS;
        private IManSession _objSession;
        private IManDatabase _objDatabase;

        public Program()
        {
            _objDMS = new ManDMS();
            _objSession = _objDMS.Sessions.Add(strServer);
            _objSession.Login(strUsername, strPassword);
            _objSession.MaxRowsForSearch = 1000;
            _objSession.MaxRowsNonSearch = 500;
            _objDatabase = _objSession.PreferredDatabase;
        }

        public IManDMS getDMS
        {
            get
            {
                return _objDMS;
            }
            set
            {
            }
        }

        public IManSession getSession
        {
            get
            {
                return _objSession;
            }
            set
            {
            }
        }

        public IManDatabase getDatabase
        {
            get
            {
                return _objDatabase;
            }
            set
            {
            }
        }

        public string do_get_folders()
        {
            IManProfileSearchParameters srch_params = _objDMS.CreateProfileSearchParameters();
            IManWorkspaceSearchParameters wkspc_srch_params = _objDMS.CreateWorkspaceSearchParameters();
            IManFolders wkspc_res;
            string folder_IDs = "";
            List <IManFolder> obj_folders = new List<IManFolder>();

            srch_params.Add(imProfileAttributeID.imProfileAuthor, "WSADMIN");
            srch_params.Add(imProfileAttributeID.imProfileCustom1, "1002279");              // BearingPoint, Inc.
            wkspc_res = _objDatabase.SearchWorkspaces(srch_params, wkspc_srch_params);

            Console.WriteLine("\nFound {0} workspaces.\n", wkspc_res.Count);

            SortWorkspaceClass objIManSortWorkspace = new SortWorkspaceClass();
            wkspc_res.Sort(objIManSortWorkspace);

            foreach (IManWorkspace work1 in wkspc_res)
                if (work1.SubType.Equals("work"))
                {
                    folder_IDs += work1.FolderID.ToString() + ",";
                    //Console.WriteLine(work1.Name);

                    foreach (IManFolder fldr1 in work1.SubFolders)
                        if (fldr1.ObjectType.ObjectType.Equals(imObjectType.imTypeDocumentFolder))
                            obj_folders.Add(fldr1);
                }

            Console.WriteLine("\nobj_folders[] contains: {0} folder_IDs\n", obj_folders.Count);

            //char[] trim_chars = { ',', ' ', '\n' };
            folder_IDs = folder_IDs.TrimEnd(',');
            return folder_IDs;
        }

        private void thread_worker(ManualResetEvent _reset_event, Mutex _mutex, string folders)
        {
            int tid = Thread.CurrentThread.ManagedThreadId;

            Thread.Sleep(200);
            // Perform 10 runs through the folder list
            for (int j = 0; j < 10; j++)
                foreach (string ID in folders.Split(','))
                {
                    IManFolder imF = _objDatabase.GetFolder(Int16.Parse(ID));
                    imF.Refresh();
                    
                    _mutex.WaitOne();
                    Console.SetCursorPosition(1, tid + 22);
                    Console.Write("Loop({0}) : GetFolder({1})\t\t", tid, ID);
                    _mutex.ReleaseMutex();
                    
                }
            _mutex.WaitOne();
            Console.WriteLine("Thread {0} exited.", tid);
            _mutex.ReleaseMutex();
            _reset_event.Set();
        }

        public void do_threads(string folder_ids)
        {
            ManualResetEvent[] _manualResetEvents = new ManualResetEvent[10];
            Mutex mutex1 = new Mutex(false);
            Stopwatch sw1 = new Stopwatch();

            sw1.Start();
            // create 10 threads
            for (int n = 0; n < 10; n++)
            {
                ManualResetEvent _mres = new ManualResetEvent(false);
                _manualResetEvents[n] = _mres;

                Thread t = new Thread(() => thread_worker(_mres, mutex1, folder_ids));
                t.Start();
                Console.WriteLine("Thread {0,2} started with ({1}, {2}, folder_ids)", t.ManagedThreadId, mutex1.ToString(), _mres.ToString());
            }
            Console.WriteLine();
            WaitHandle.WaitAll(_manualResetEvents);
            sw1.Stop();
            Console.WriteLine("\n\n\n\n--------------------------------------------------");
            Console.WriteLine("\ndo_threads: {0} ms", sw1.ElapsedMilliseconds);
        }

        static void display_exception_details(Exception e)
        {
            Console.WriteLine("Message:\n{0{\n\nSource:\n{1}\n\nStack Trace:\n{2}\n\n",
                                e.Message, e.Source, e.StackTrace);
            Console.Write("Press Enter: ");
            Console.ReadLine();
            System.Environment.Exit(-1);
        }

        static void Main(string[] args)
        {
            Stopwatch sw1 = new Stopwatch();
            sw1.Start();
            Program prog = new Program();
            sw1.Stop();
            Console.WriteLine("\nprog.Program: {0} ms\n", sw1.ElapsedMilliseconds);

            Console.WriteLine("_objDMS.ComputerName\t\t= {0}", prog.getDMS.ComputerName);
            Console.WriteLine("_objSession.Timeout\t\t= {0}", prog.getSession.Timeout);
            Console.WriteLine("_objSession.AllVersions\t\t= {0}", prog.getSession.AllVersions);
            Console.WriteLine("_objSession.MaxRowsForSearch\t= {0}", prog.getSession.MaxRowsForSearch);

            string str1 = prog.do_get_folders();
            prog.do_threads(str1);

            if (prog._objSession.Connected)
                try
                {
                    sw1.Restart();
                    prog._objSession.Logout();
                    sw1.Stop();
                    Console.WriteLine("\n_objSession.Logout: {0} ms", sw1.ElapsedMilliseconds);
                }
                catch (COMException e)
                {
                    display_exception_details(e);
                }
 
            Console.Write("\nPress Enter: ");
            Console.Beep();
            Console.ReadLine();
        }
    }
}
