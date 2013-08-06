using System;
using System.Windows;
using System.Windows.Input;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Diagnostics;
using AForge.Video.FFMPEG;
using System.Runtime.InteropServices;

namespace Screen_Capture_Tool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Application Variables

        private Thread captureThread; //The thread object for capturing asynchronously
        private bool requestThreadShutdown; //A variable which is used to allow clean stopping of the capture thread and determine if the escape key should close the application

        private const String outputFilename = "output.mp4"; //The output filename
        private String imageDirPath; //The path to where the output is stored
        private String appDir; //The path of this executable
        
        private const int captureInterval = 33; //The interval between captures in ms
        private const int maxCaptureLength = 216000000; //The maximum number of captures in ms
        private const int frameRateScaleFactor = 1000; //1000 represents milliseconds as frameRate=frameRateScaleFactor/captureInterval
        private const int bitRate = 100000000; //Bigger means bigger files and less compression (better quality)

        private const int cursorOffsetLeft = 4; //Cursor left offset for drawing
        private const int cursorOffsetTop = 0; //Cursor top offset for drawing
        private System.Windows.Media.Brush backgroundBrush; //Variable to store the initial state of the background for reverting after recording
        
        #endregion

        #region Mouse Capture Code

        [StructLayout(LayoutKind.Sequential)]
        struct CURSORINFO
        {
            public Int32 cbSize;
            public Int32 flags;
            public IntPtr hCursor;
            public POINTAPI ptScreenPos;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct POINTAPI
        {
            public int x;
            public int y;
        }

        [DllImport("user32.dll")]
        static extern bool GetCursorInfo(out CURSORINFO pci);

        [DllImport("user32.dll")]
        static extern bool DrawIconEx(IntPtr hDC, int X, int Y, IntPtr hIcon, int cxWidth, int cyWidth, uint istepIfAniCur, IntPtr hbrFlickerFreeDraw, uint diFlags);

        const Int32 CURSOR_SHOWING = 0x00000001;
        const Int32 DI_COMPAT = 0x0004;
        const Int32 DI_DEFAULTSIZE = 0x0008;
        const Int32 DI_IMAGE = 0x0002;
        const Int32 DI_MASK = 0x0001;
        const Int32 DI_NOMIRROR = 0x0010;
        const Int32 DI_NORMAL = 0x0003;

        #endregion

        #region Constructor

        public MainWindow()
        {
            InitializeComponent();

            //Register Important Window Events
            this.MouseDown+=new MouseButtonEventHandler(Window_MouseDown); //For dragging the form
            this.KeyDown += new KeyEventHandler(MainWindow_KeyDown); //For keyboard shortcuts
            this.SizeChanged += new SizeChangedEventHandler(MainWindow_SizeChanged); //For modifying the capture region size

            //Generate File Constants
            imageDirPath = "";
            String[] pathParts = System.Reflection.Assembly.GetExecutingAssembly().Location.Split(new char[] { '\\' });
            for (int i = 0; i < pathParts.Length - 1; i++) //Strip out the exe name
                imageDirPath += pathParts[i] + '\\';
            appDir = imageDirPath; //Store the app path
            imageDirPath += "Output\\"; //Store the output directory path

            //Create Output Dir
            Directory.CreateDirectory(imageDirPath);

            //Store the initial background brush
            backgroundBrush = this.Background.Clone();

            //Show instructions
            notificationPopupText.Text = "Resize, drag, record for up to 6 hours. Escape to stop recording or quit if you're not recording. F1 or button to record.\r\n<Click to Close...>";
            notificationPopup.IsOpen = true;

            //Sync Popup Location to Window
            this.LocationChanged += delegate(object sender, EventArgs args)
            {
                var offset = notificationPopup.HorizontalOffset;
                notificationPopup.HorizontalOffset = offset + 1;
                notificationPopup.HorizontalOffset = offset;
            };
        }

        #endregion

        #region Window Events

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) //Left click drags the form
                this.DragMove();
        }

        void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) //Escape closes the form and attempts to end the capture process
            {
                if (requestThreadShutdown==false&&captureThread!=null)
                    requestThreadShutdown = true;
                else
                    this.Close();
            }
            else if (e.Key == Key.F1)
                recordButton_Click(null, null);
        }

        void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            //Update the border size (all other updates are implicit)
            windowBorder.Width = e.NewSize.Width;
            windowBorder.Height = e.NewSize.Height;
        }

        #endregion

        #region Async GUI Callbacks

        private delegate void ReEnableFormCallback();
        private void ReEnableForm()
        {
            //Set the form in an enabled state by performing the following
            recordButton.Visibility = System.Windows.Visibility.Visible; //Make the button visible
            this.Background = backgroundBrush; //Make the background the default background brush color
            this.ResizeMode = System.Windows.ResizeMode.CanResizeWithGrip; //Reenable the grip for resize
        }

        private delegate object GetParameterUpdateCallback();
        private object GetParameterUpdate()
        {
            //Return a dynamic structure containing important async thread state information
            return new
            {
                Left = (int)this.Left,
                Top = (int)this.Top
            };
        }

        private delegate void ShowPopupCallback(String popupText);
        private void ShowPopup(String popupText)
        {
            //Show the popup with the appropriate text
            notificationPopupText.Text = popupText;
            notificationPopup.IsOpen = true;
            notificationPopup.StaysOpen = true;
        }

        #endregion

        #region Control Events

        private void recordButton_Click(object sender, RoutedEventArgs e)
        {
            //When the record button is pressed, place the GUI in a recording state
            recordButton.Visibility = System.Windows.Visibility.Hidden; //Make the button hidden
            this.Background = System.Windows.Media.Brushes.Transparent; //Make the background transparent
            this.ResizeMode = System.Windows.ResizeMode.NoResize; //Make the form fixed size
            notificationPopup.Visibility = System.Windows.Visibility.Hidden;
            notificationPopup.IsOpen = false;
            captureThread = new Thread(new ParameterizedThreadStart(captureLoop)); //Instantiate the capture thread
            requestThreadShutdown = false;
            captureThread.Start(new { BorderThickness = windowBorder.BorderThickness, Left = (int)this.Left, Top = (int)this.Top, Width = (int)this.Width, Height = (int)this.Height, Interval = captureInterval, Iterations = maxCaptureLength }); //Start the capture thread with initial parameters
        }

        private void notificationPopup_MouseDown(object sender, MouseButtonEventArgs e)
        {
            //Close the popup when it is clicked
            notificationPopup.PopupAnimation = System.Windows.Controls.Primitives.PopupAnimation.Fade;
            notificationPopup.StaysOpen = false;
            notificationPopup.IsOpen = false;
        }

        #endregion

        #region Async Capture Thread

        private void captureLoop(object parameters)
        {
            dynamic p = parameters;

            //Precompute size parameters s.t. actual width/height are multiples of two and requested width/height are the window size
            Thickness borderThickness = p.BorderThickness;
            int bLeft = (int)borderThickness.Left;
            int bRight = (int)borderThickness.Right;
            int bTop = (int)borderThickness.Top;
            int bBottom = (int)borderThickness.Bottom;
            int requestedWidth = p.Width - bLeft - bRight;
            int requestedHeight = p.Height - bTop - bBottom;
            int actualWidth = requestedWidth + requestedWidth % 2; //Enforce multiple of two
            int actualHeight = requestedHeight + requestedHeight % 2; //Enforce multiple of two
            int left = p.Left;
            int top = p.Top;
            System.Drawing.Size requestedSize = new System.Drawing.Size(requestedWidth, requestedHeight);

            //Generate the render surfaces
            Bitmap requestedRenderBitmap = new Bitmap(requestedWidth, requestedHeight);
            Graphics requestedRenderGraphics = Graphics.FromImage(requestedRenderBitmap);
            Bitmap actualRenderBitmap = new Bitmap(actualWidth, actualHeight);
            Graphics actualRenderGraphics = Graphics.FromImage(actualRenderBitmap);

            //Create a timer for keeping loop time as consistent as possible
            Stopwatch loopTimer = new Stopwatch(); //Timer for timing each iteration of the loop (for convenience)
            Stopwatch fullCaptureTimer = new Stopwatch(); //Timer for timing the full capture
            int performanceCounter = 0; //Create a local variable for tracking performance in number of ms per loop interval missed
            int maxFrameRateDeviation = 0; //Create a local variable for tracking the maximum frame rate deviation
            int deviationCount = 0; //Count the number of frames which did not successfully meet their timing constraint
            int framesWritten = 0; //Count the number of frames which have been written to file

            //Create output stream
            VideoFileWriter writer = new VideoFileWriter();
            writer.Open(imageDirPath + outputFilename, actualWidth, actualHeight, frameRateScaleFactor / captureInterval, VideoCodec.MPEG4, bitRate); //Open mp4 file using multiple of two width and height

            //Start the full capture timer
            fullCaptureTimer.Start();

            //Begin capturing
            int i;
            for (i = 0; i < maxCaptureLength; i++)
            {
                loopTimer.Restart(); //Restart the timer for loop speed regulating

                //Copy the screen
                requestedRenderGraphics.CopyFromScreen(left + bLeft, top + bTop, 0, 0, requestedSize);

                //Draw the cursor on the image
                CURSORINFO pci;
                pci.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(CURSORINFO));
                if (GetCursorInfo(out pci))
                {
                    if (pci.flags == CURSOR_SHOWING)
                    {
                        DrawIconEx(requestedRenderGraphics.GetHdc(),
                                    pci.ptScreenPos.x - left - bLeft - cursorOffsetLeft,
                                    pci.ptScreenPos.y - top - bTop - cursorOffsetTop,
                                    pci.hCursor,
                                    0, 0,
                                    0, IntPtr.Zero,
                                    DI_MASK | DI_IMAGE | DI_COMPAT);
                        requestedRenderGraphics.ReleaseHdc();
                    }
                }

                //Output to Video
                actualRenderGraphics.DrawImage(requestedRenderBitmap, 0, 0); //Render the graphics down to the output surface
                writer.WriteVideoFrame(actualRenderBitmap); //Write the frame to the output stream
                framesWritten++;

                p = this.Dispatcher.Invoke(new GetParameterUpdateCallback(GetParameterUpdate), null); //Update the system parameters (only position is used)
                if (p != null)
                {
                    //Update the movement variables if they're valid
                    left = p.Left;
                    top = p.Top;
                }

                loopTimer.Stop(); //Stop the timer

                //If we've gone over time, increment the performance counter by the number of ms we missed by
                int deviation = (long)captureInterval < loopTimer.ElapsedMilliseconds ?
                    (int)(loopTimer.ElapsedMilliseconds - (long)captureInterval) 
                    : 0;
                if (deviation > 0) deviationCount++;
                if (deviation > maxFrameRateDeviation) maxFrameRateDeviation = deviation; //Set the new max if appropriate
                performanceCounter += deviation;

                //Computer makeup frames and duplicate write as many times as is necessary
                int expectedFrames = (int)((double)fullCaptureTimer.ElapsedMilliseconds / (double)captureInterval);
                for (int frameCount = framesWritten; framesWritten<expectedFrames; framesWritten++) //For as many times as we can subtract the frame interval from the makeup counter (meaning how many frames we missed by taking too long)
                    writer.WriteVideoFrame(actualRenderBitmap); //Write the frame to the output stream

                //Watch for shutdown condition
                if (requestThreadShutdown)
                {
                    //Stop the full capture timer to see how well we did
                    fullCaptureTimer.Stop();

                    //Exit the capture loop
                    break;
                }

                //If we're under time, sleep until it's time to start again
                Thread.Sleep((long)captureInterval > loopTimer.ElapsedMilliseconds ?
                    (int)((long)captureInterval - loopTimer.ElapsedMilliseconds) 
                    : 0);
            }
            
            //Close the output stream
            writer.Close();

            //Reenable the form
            recordButton.Dispatcher.Invoke(new ReEnableFormCallback(ReEnableForm), null);

            //Show performance
            this.Dispatcher.Invoke(new ShowPopupCallback(ShowPopup), new object[] { "Finished with " + 
                (performanceCounter / i) + "ms average framerate deviation, " + deviationCount + 
                "/" + (i + 1) + " total missed frames, and " + maxFrameRateDeviation + 
                "ms maximum frame deviation.\r\nFound "+framesWritten+" frames; expected "+(fullCaptureTimer.ElapsedMilliseconds/captureInterval)+".\r\n<Click to Close...>" });
        }

        #endregion
    }
}
