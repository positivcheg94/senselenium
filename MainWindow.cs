using System;
using System.Windows;
using System.IO;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;

using System.Diagnostics;

namespace Senselenium
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    
    class WrappedLong {
        public long i;
        public WrappedLong(long x) {i=x;}
    }
    
    public partial class MainWindow
    {        
        PXCMSenseManager psm;
        PXCMTouchlessController ptc;

        const double scrollSensitivity = 10f;

        IWebDriver browser;

        Actions browserActs;

        double zoomDelta = 0.0;
        double prevZoomZ = 0.0;

        double scrollXDelta = 0.0;
        double prevScrollX = 0.0;
        double scrollYDelta = 0.0;
        double prevScrollY = 0.0;

        WrappedLong swipeLeftLast = new WrappedLong(0);
        WrappedLong swipeRightLast = new WrappedLong(0);
        WrappedLong swipeUpLast = new WrappedLong(0);
        WrappedLong swipeDownLast = new WrappedLong(0);
        WrappedLong openTabLast = new WrappedLong(0);
        WrappedLong closeTabLast = new WrappedLong(0);
        WrappedLong prevTabLast = new WrappedLong(0);


        Stopwatch stopwatch;

        public MainWindow(string s)
        {
            stopwatch = new Stopwatch();
            stopwatch.Start();

            if (s.Length > 0)
            {
                try
                {
                    if (s.ToLower()[0] == 'f' || Path.GetFileNameWithoutExtension(s).ToLower()[0] == 'f')
                    {
                        browser = new FirefoxDriver();
                    }
                    else
                    {
                        browser = new ChromeDriver();
                    }
                }
                catch
                {
                    browser = new ChromeDriver();
                }
            }
            else
            {
                Console.WriteLine("Open Firefox(F) or Chrome(C)?");
                string choice = Console.ReadLine();
                if (choice.ToUpper() == "F")
                {
                    browser = new FirefoxDriver();
                }
                else
                {
                    browser = new ChromeDriver();
                }
            }
            Window_Loaded(null, null);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            browserActs = new Actions(browser);
            browser.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(5));

            try
            {
                browser.Navigate().GoToUrl("http://www.google.com/");
            }
            catch { }

            StartRealSense();

            UpdateConfiguration();

            StartFrameLoop();

            MessageBox.Show("Режим Senselenium активирован! \r\n Теперь вы можете испытать браузинг нового поколения!","Загрузка завершена",MessageBoxButton.OK,MessageBoxImage.Information,MessageBoxResult.OK,MessageBoxOptions.ServiceNotification);
 
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StopRealSense();

            browser.Quit();

            stopwatch.Stop();
        }

        private void Log(String s)
        {
            Console.WriteLine("[" + stopwatch.ElapsedMilliseconds + "] " + s);
        }


        private void StartRealSense()
        {
            Log("Starting Touchless Controller");

            pxcmStatus rc;

            // creating Sense Manager
            psm = PXCMSenseManager.CreateInstance();
            Log("Creating SenseManager: " + psm == null ? "failed" : "success");
            if (psm == null)
                Environment.Exit(-1);

            // Optional steps to send feedback to Intel Corporation to understand how often each SDK sample is used.
            PXCMMetadata md = psm.QuerySession().QueryInstance<PXCMMetadata>();
            if (md != null)
            {
                string sample_name = "Touchless Listbox CS";
                md.AttachBuffer(1297303632, System.Text.Encoding.Unicode.GetBytes(sample_name));
            }

            // work from file if a filename is given as command line argument
            string[] args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                psm.captureManager.SetFileName(args[1], false);
            }

            // Enable touchless controller in the multimodal pipeline
            rc = psm.EnableTouchlessController(null);
            Log("Enabling Touchless Controller: " + rc.ToString());
            if (rc != pxcmStatus.PXCM_STATUS_NO_ERROR)
                Environment.Exit(-1);

            // initialize the pipeline
            PXCMSenseManager.Handler handler = new PXCMSenseManager.Handler();
            rc = psm.Init(handler);
            Log("Initializing the pipeline: " + rc.ToString());
            if (rc != pxcmStatus.PXCM_STATUS_NO_ERROR)
                Environment.Exit(-1);

            // getting touchless controller
            ptc = psm.QueryTouchlessController();
            if (ptc == null)
                Environment.Exit(-1);
            ptc.SubscribeEvent(new PXCMTouchlessController.OnFiredUXEventDelegate(OnTouchlessControllerUXEvent));

            ptc.AddGestureActionMapping("swipe_left", PXCMTouchlessController.Action.Action_LeftKeyPress, OnFiredAction);
            ptc.AddGestureActionMapping("swipe_right", PXCMTouchlessController.Action.Action_RightKeyPress, OnFiredAction);
            ptc.AddGestureActionMapping("swipe_up", PXCMTouchlessController.Action.Action_VolumeUp, OnFiredAction);
            ptc.AddGestureActionMapping("swipe_down", PXCMTouchlessController.Action.Action_VolumeDown, OnFiredAction);
            ptc.AddGestureActionMapping("thumb_up", PXCMTouchlessController.Action.Action_PlayPause, OnFiredAction);
            //ptc.AddGestureActionMapping("thumb_down", PXCMTouchlessController.Action.Action_Stop, OnFiredAction);
            ptc.AddGestureActionMapping("two_fingers_pinch_open", PXCMTouchlessController.Action.Action_Mute, OnFiredAction);
            //ptc.AddGestureActionMapping("fist", PXCMTouchlessController.Action.Action_PrevTrack, OnFiredAction);
        }

        // on closing
        private void StopRealSense()
        {
            Log("Disposing SenseManager and Touchless Controller");
            ptc.Dispose();
            psm.Close();
            psm.Dispose();
           
        }
        
        private void UpdateConfiguration()
        {
            pxcmStatus rc;
            PXCMTouchlessController.ProfileInfo pInfo;

            rc = ptc.QueryProfile(out pInfo);
            Log("Querying Profile: " + rc.ToString());
            if (rc != pxcmStatus.PXCM_STATUS_NO_ERROR)
                Environment.Exit(-1);

            pInfo.config = PXCMTouchlessController.ProfileInfo.Configuration.Configuration_Scroll_Vertically |
                PXCMTouchlessController.ProfileInfo.Configuration.Configuration_Scroll_Horizontally|
                PXCMTouchlessController.ProfileInfo.Configuration.Configuration_Allow_Selection|
                PXCMTouchlessController.ProfileInfo.Configuration.Configuration_Allow_Zoom|
                PXCMTouchlessController.ProfileInfo.Configuration.Configuration_Meta_Context_Menu
                |PXCMTouchlessController.ProfileInfo.Configuration.Configuration_Allow_Back
                ;

            rc = ptc.SetProfile(pInfo);
            Log("Setting Profile: " + rc.ToString());
        }

        private void StartFrameLoop()
        {
            psm.StreamFrames(false);
        }

        private void OnTouchlessControllerUXEvent(PXCMTouchlessController.UXEventData data)
        {
                switch (data.type)
                {
                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_CursorVisible:
                        {
                            Log("Cursor Visible");
                        }
                        break;
                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_CursorNotVisible:
                        {
                            Log("Cursor Not Visible");
                        }
                        break;
                    /*case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_GotoStart:
                        {
                            browser.Navigate().Refresh();
                            Log("Nice wave!");
                        }   
                        break;*/
                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_Select:
                        {
                            if (browser.WindowHandles.Count > 1)
                            {
                                lock (closeTabLast)
                                {
                                    if (stopwatch.ElapsedMilliseconds - closeTabLast.i >= 1500)
                                    {

                                        browserActs.KeyDown(Keys.Control).SendKeys("w").KeyUp(Keys.Control).Build().Perform();
                                        Log("Close tab!!!");    
                                        closeTabLast.i = stopwatch.ElapsedMilliseconds;
                                    }
                                }
                            }
                        }
                        break;
                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_StartScroll:
                        {
                            Log("Start Scroll");
                            prevScrollX = data.position.x;
                            prevScrollY = data.position.y;
                        }
                        break;
                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_EndScroll:
                        {
                            Log("End Scroll");
                            prevScrollX = data.position.y;
                            scrollXDelta = 0.0;
                            prevScrollY =  data.position.y;
                            scrollYDelta= 0.0;
                        }
                        break;
                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_CursorMove:
                        {

                            Point point = new Point();
                            point.X = Math.Max(Math.Min(0.99F, data.position.x), 0.01F);
                            point.Y = Math.Max(Math.Min(0.99F, data.position.y), 0.01F);

                            try
                            {
                                int mouseX = (int)(browser.Manage().Window.Position.X + point.X * browser.Manage().Window.Size.Width);
                                int mouseY = (int)(browser.Manage().Window.Position.Y + point.Y * browser.Manage().Window.Size.Height);
                                MouseInjection.SetCursorPos(mouseX, mouseY);
                            }
                            catch { }

                            
                            
                        }
                        break;
                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_Scroll:
                        {
                            //myListscrollViwer.ScrollToVerticalOffset(initialScrollOffest + (data.position.y - initialScrollPoint) * scrollSensitivity);
                            const double scrollThreshold = 0.005;
                            scrollXDelta = data.position.x - prevScrollX;
                            scrollYDelta = data.position.y - prevScrollY;
                            prevScrollX = data.position.x;
                            prevScrollY = data.position.y;
                            if (scrollXDelta > 0)
                            {
                                while (scrollXDelta >= scrollThreshold)
                                {

                                    browserActs.SendKeys(Keys.ArrowRight).Build().Perform();
                                    scrollXDelta -= scrollThreshold;
                                }

                            }
                            else if (scrollXDelta < 0)
                            {
                                while (scrollXDelta <= -scrollThreshold)
                                {
                                    browserActs.SendKeys(Keys.ArrowLeft).Build().Perform();
                                    scrollXDelta += scrollThreshold;
                                }

                            }
                            if (scrollYDelta > 0)
                            {
                                while (scrollYDelta >= scrollThreshold) { 
              
                                    browserActs.SendKeys(Keys.ArrowDown).Build().Perform();
                                    scrollYDelta -= scrollThreshold;
                                }
                              
                            }
                            else if (scrollYDelta < 0)
                            {
                                while (scrollYDelta <= -scrollThreshold)
                                {
                                    browserActs.SendKeys(Keys.ArrowUp).Build().Perform();
                                    scrollYDelta += scrollThreshold;
                                }
                           
                            }

                            Log("Scrolling: dx="+scrollXDelta+" dy="+scrollYDelta);
                            scrollXDelta = 0.0;
                            scrollYDelta = 0.0;
                        }
                        break;
                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_StartZoom:

                        Log(String.Format("Gesture: Start Zoom"));
                        prevZoomZ = data.position.z;
                        break;
                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_Zoom:
                        zoomDelta = prevZoomZ - data.position.z;
                         if (Math.Abs(zoomDelta) > 0.05)
                        {
                           prevZoomZ = data.position.z;
                           Log("Gesture: Zooming by "+zoomDelta);
                           if (zoomDelta > 0)
                           {
                               browserActs.KeyDown(Keys.Control).SendKeys("=").KeyUp(Keys.Control).Build().Perform();
                           }
                           else
                           {
                               browserActs.KeyDown(Keys.Control).SendKeys("-").KeyUp(Keys.Control).Build().Perform();

                           }
                            
                        }
                        break;
                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_EndZoom:
                        Log("Gesture: End Zoom");
                        break;
                   /* case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_Back:
                        Log("Gesture: Back");
                        try
                        {
                            browser.Navigate().Back();
                        }
                        catch { }
                        break;
                    */
                   case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_ShowMetaMenu:
               
                        break;

                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_HideMetaMenu:
                    
                        break;

                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_MetaOpenHand:
                        {
                            Log("Refresh");
                            try
                            {
                                browser.Navigate().Refresh();
                            }
                            catch { }
                        }
                        break;

                    case PXCMTouchlessController.UXEventData.UXEventType.UXEvent_MetaPinch:
                        {
                            Log("Maximize");
                            browser.Manage().Window.Maximize();
                        }
                        break;
                }
        }

        public void OnFiredAction(PXCMTouchlessController.Action action)
        {
         
            
            switch (action)
            {
                
                case PXCMTouchlessController.Action.Action_LeftKeyPress:
                    {
                        lock(swipeLeftLast)
                        {
                            if (stopwatch.ElapsedMilliseconds - swipeLeftLast.i >= 500)
                            {
                                browserActs.KeyDown(Keys.Control).KeyDown(Keys.Shift).SendKeys(Keys.Tab).KeyUp(Keys.Shift).KeyUp(Keys.Control).Build().Perform();
                                Log("Left Swipe");
                                swipeLeftLast.i = stopwatch.ElapsedMilliseconds;
                            }
                        }
                    }
                break;
                case PXCMTouchlessController.Action.Action_RightKeyPress:
                {
                    lock(swipeRightLast)
                        {
                            if (stopwatch.ElapsedMilliseconds - swipeRightLast.i >= 500)
                            {
                                browserActs.KeyDown(Keys.Control).SendKeys(Keys.Tab).KeyUp(Keys.Control).Build().Perform();
                                Log("Right Swipe");
                                swipeRightLast.i = stopwatch.ElapsedMilliseconds;
                            }
                        }
                }
                break;
                case PXCMTouchlessController.Action.Action_VolumeUp:
                {
                    lock (swipeUpLast)
                    {
                        if (stopwatch.ElapsedMilliseconds - swipeUpLast.i >= 500)
                            {
                                Log("Gesture: Back");
                                try
                                {
                                    browserActs.KeyDown(Keys.Alt).SendKeys(Keys.Left).KeyUp(Keys.Alt);
                                }
                                catch { }
                                swipeUpLast.i = stopwatch.ElapsedMilliseconds;
                            }
                        }
                }
                break;
                case PXCMTouchlessController.Action.Action_VolumeDown:
                {
                    lock (swipeDownLast)
                    {
                        if (stopwatch.ElapsedMilliseconds - swipeDownLast.i >= 500)
                            {
                                try
                                {
                                    browserActs.KeyDown(Keys.Alt).SendKeys(Keys.Right).KeyUp(Keys.Alt);
                                }
                                catch { }
                                Log("Down Swipe");
                                swipeDownLast.i = stopwatch.ElapsedMilliseconds;
                            }
                    }
                }

                break;
               case PXCMTouchlessController.Action.Action_PlayPause:
                {
                    lock (openTabLast)
                    {
                        if (stopwatch.ElapsedMilliseconds - openTabLast.i >= 1500)
                            {
                                browserActs.KeyDown(Keys.Control).SendKeys("t").KeyUp(Keys.Control).Build().Perform();
                                Log("open tab!!!");
                                openTabLast.i = stopwatch.ElapsedMilliseconds;
                            }
                    }
                }
                break;
                case PXCMTouchlessController.Action.Action_Mute:
                {
                    Log("Select");
                    MouseInjection.ClickLeftMouseButton();
                }
                break;
                /*case PXCMTouchlessController.Action.Action_PrevTrack:
                {
                    lock (prevTabLast)
                    {
                        if (stopwatch.ElapsedMilliseconds - prevTabLast.i >= 2500)
                        {

                            browserActs.KeyDown(Keys.Control).KeyDown(Keys.Shift).SendKeys("t").KeyUp(Keys.Shift).KeyUp(Keys.Control).Build().Perform();
                            Log("Previous tab!!!");
                            prevTabLast.i = stopwatch.ElapsedMilliseconds;
                        }
                    }
                }
                break;*/
            }
        }
    }
}
