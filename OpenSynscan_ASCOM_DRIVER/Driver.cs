// --------------------------------------------------------------------------------
//
// ASCOM Telescope driver for OpenSynscan
//
// Implements:	ASCOM Telescope interface
// Author:		Vicente Sirvent <vicentesirvent@gmail.com
//
// Edit Log:
//
// Date			Who	Vers	Description
// -----------	---	-----	-------------------------------------------------------
// 30/09/2019	VSO	1.0.0	Initial edit, created from ASCOM driver template
// --------------------------------------------------------------------------------
//


#define Telescope

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Runtime.InteropServices;

using ASCOM;
using ASCOM.Astrometry;
using ASCOM.Astrometry.AstroUtils;
using ASCOM.Utilities;
using ASCOM.DeviceInterface;
using System.Globalization;
using System.Collections;
using System.Net.Sockets;
using System.Net;
using System.Threading;

namespace ASCOM.OpenSynscan
{
    /// <summary>
    /// ASCOM Telescope Driver for OpenSynscan.
    /// </summary>
    [Guid("9ffb6dd5-1cea-4775-92ed-e09565f870b7")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Telescope : ITelescopeV3
    {
        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>
        internal static string driverID = "ASCOM.OpenSynscan.Telescope";
        // TODO Change the descriptive string for your driver then remove this line
        /// <summary>
        /// Driver description that displays in the ASCOM Chooser.
        /// </summary>
        private static string driverDescription = "Pulse Guide for OpenSynscan 1.0";
        /// <summary>
        /// Private variable to hold the connected state
        /// </summary>
        private bool connectedState;

        /// <summary>
        /// Private variable to hold an ASCOM Utilities object
        /// </summary>
        private Util utilities;

        /// <summary>
        /// Private variable to hold an ASCOM AstroUtilities object to provide the Range method
        /// </summary>
        private AstroUtils astroUtilities;

        /// <summary>
        /// Variable to hold the trace logger object (creates a diagnostic log file with information that you specify)
        /// </summary>
        internal static TraceLogger tl;

        public const int TX_PORT = 5002;
        public const int RX_PORT = 5003;
        public const int TX_PKT_LEN = 8;
        public const int RX_PKT_LEN = 2;
        public const byte PULSE_PKT_START = 0x35;
        // 0 (1 byte) - PULSE_PKT_START
        // 1 (1 byte) - Sequence
        // 2 (1 byte) - DIRECTION :
        //     +DEC = 0,
        //     -DEC = 1,
        //     +RA  = 2,
        //     -RA  = 3
        // 3 (4 bytes) - msec.
        // 7 (1 byte) - rate x 10 
        public const byte SYNSCAN_PKT_START = 0x36;
        private UdpClient mSender;
        private int mSeq = 1;
        private UdpClient mReceiver;
        private IPEndPoint mDevice;
        private Byte[] mTxPkt;
        private Byte[] mRxPkt;
        private SocketAsyncEventArgs mRxEvent;
        private Mutex mMutex;
        private DateTime mLastCommandEndTime = DateTime.Now;
        /// <summary>
        /// Initializes a new instance of the <see cref="OpenSynscan"/> class.
        /// Must be public for COM registration.
        /// </summary>
        public Telescope()
        {
            tl = new TraceLogger("", "OpenSynscan");
            ReadProfile(); // Read device configuration from the ASCOM Profile store

            tl.LogMessage("Telescope", "Starting initialisation");

            connectedState = false; // Initialise connected to false
            utilities = new Util(); //Initialise util object
            astroUtilities = new AstroUtils(); // Initialise astro utilities object
            mMutex = new Mutex();
            mSender = new UdpClient(TX_PORT);
            mReceiver = new UdpClient(RX_PORT);

            mDevice = null;
            mTxPkt = new Byte[TX_PKT_LEN];
            mRxPkt = new Byte[RX_PKT_LEN];
            mTxPkt[0] = PULSE_PKT_START;
            mRxEvent = new SocketAsyncEventArgs();
            mRxEvent.SetBuffer(mRxPkt, 0, RX_PKT_LEN);
            mRxEvent.RemoteEndPoint = new IPEndPoint(IPAddress.Any, 0);
            mRxEvent.UserToken = mReceiver.Client;
            mRxEvent.Completed += new EventHandler<SocketAsyncEventArgs>(RecvCompleted);
            mReceiver.Client.ReceiveFromAsync(mRxEvent);

            tl.LogMessage("Telescope", "Completed initialisation");
        }

        void RecvCompleted(object sender, SocketAsyncEventArgs e)
        {
            mMutex.WaitOne();
            try
            {
                byte pt = mRxPkt[0];
                if (mDevice == null && pt == SYNSCAN_PKT_START)
                {
                    tl.LogMessage("Telescope", "Opensynscan discovered at " + e.RemoteEndPoint.ToString());
                    mDevice = (IPEndPoint) e.RemoteEndPoint;
                    mSender.Client.Connect(mDevice);
                }
                else
                {
                    mReceiver.Client.ReceiveFromAsync(mRxEvent);
                }
            }
            catch (Exception) { }
            mMutex.ReleaseMutex();
        }

        //
        // PUBLIC COM INTERFACE ITelescopeV3 IMPLEMENTATION
        //

        #region Common properties and methods.

        /// <summary>
        /// Displays the Setup Dialog form.
        /// If the user clicks the OK button to dismiss the form, then
        /// the new settings are saved, otherwise the old values are reloaded.
        /// THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        /// </summary>
        public void SetupDialog()
        {
            // consider only showing the setup dialog if not connected
            // or call a different dialog if connected
            if (IsConnected)
            {
                System.Windows.Forms.MessageBox.Show("Already connected, just press OK");
            }

            using (SetupDialogForm F = new SetupDialogForm())
            {
                var result = F.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
                }
            }
        }

        public ArrayList SupportedActions
        {
            get
            {
                tl.LogMessage("SupportedActions Get", "Returning empty arraylist");
                return new ArrayList();
            }
        }

        public string Action(string actionName, string actionParameters)
        {
            LogMessage("", "Action {0}, parameters {1} not implemented", actionName, actionParameters);
            throw new ASCOM.ActionNotImplementedException("Action " + actionName + " is not implemented by this driver");
        }

        public void CommandBlind(string command, bool raw)
        {
            CheckConnected("CommandBlind");
            // Call CommandString and return as soon as it finishes
            this.CommandString(command, raw);
            // or
            throw new ASCOM.MethodNotImplementedException("CommandBlind");
            // DO NOT have both these sections!  One or the other
        }

        public bool CommandBool(string command, bool raw)
        {
            CheckConnected("CommandBool");
            string ret = CommandString(command, raw);
            // TODO decode the return string and return true or false
            // or
            throw new ASCOM.MethodNotImplementedException("CommandBool");
            // DO NOT have both these sections!  One or the other
        }

        public string CommandString(string command, bool raw)
        {
            CheckConnected("CommandString");
            // it's a good idea to put all the low level communication with the device here,
            // then all communication calls this function
            // you need something to ensure that only one command is in progress at a time

            throw new ASCOM.MethodNotImplementedException("CommandString");
        }

        public void Dispose()
        {
            // Clean up the tracelogger and util objects
            tl.Enabled = false;
            tl.Dispose();
            tl = null;
            utilities.Dispose();
            utilities = null;
            astroUtilities.Dispose();
            astroUtilities = null;
        }

        public bool Connected
        {
            get
            {
                LogMessage("Connected", "Get {0}", IsConnected);
                return IsConnected;
            }
            set
            {
                tl.LogMessage("Connected", "Set {0}", value);
                if (value == IsConnected)
                    return;

                if (value)
                {
                    connectedState = true;
                    LogMessage("Connected Set", "Connecting to telescope");
                    // TODO connect to the device
                }
                else
                {
                    connectedState = false;
                    LogMessage("Connected Set", "Disconnecting from telescope");
                    mDevice = null;
                }
            }
        }

        public string Description
        {
            // TODO customise this device description
            get
            {
                tl.LogMessage("Description Get", driverDescription);
                return driverDescription;
            }
        }

        public string DriverInfo
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                // TODO customise this driver description
                string driverInfo = "Information about the driver itself. Version: " + String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        public string DriverVersion
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        public short InterfaceVersion
        {
            // set by the driver wizard
            get
            {
                LogMessage("InterfaceVersion Get", "3");
                return Convert.ToInt16("3");
            }
        }

        public string Name
        {
            get
            {
                string name = "OpenSynscan";
                tl.LogMessage("Name Get", name);
                return name;
            }
        }

        #endregion

        #region ITelescope Implementation

        //////////////////////////////////////////////
        //PULSE GUIDE METHODS
        //////////////////////////////////////////////

        //
        // Resumen:
        //     Moves the scope in the given direction for the given interval or time at
        //     the rate given by the corresponding guide rate property
        //
        // Parámetros:
        //   Direction:
        //     The direction in which the guide-rate motion is to be made
        //
        //   Duration:
        //     The duration of the guide-rate motion (milliseconds)
        //
        // Excepciones:
        //   ASCOM.MethodNotImplementedException:
        //     If the method is not implemented and ASCOM.DeviceInterface.ITelescopeV3.CanPulseGuide
        //     is False
        //
        //   ASCOM.InvalidValueException:
        //     If an invalid direction or duration is given.
        //
        // Comentarios:
        //     This method returns immediately if the hardware is capable of back-to-back
        //     moves, i.e. dual-axis moves. For hardware not having the dual-axis capability,
        //     the method returns only after the move has completed.
        //     NOTES: Raises an error if ASCOM.DeviceInterface.ITelescopeV3.AtPark is true.
        //      The ASCOM.DeviceInterface.ITelescopeV3.IsPulseGuiding property must be be
        //     True during pulse-guiding.  The rate of motion for movements about the right
        //     ascension axis is specified by the ASCOM.DeviceInterface.ITelescopeV3.GuideRateRightAscension
        //     property. The rate of motion for movements about the declination axis is
        //     specified by the ASCOM.DeviceInterface.ITelescopeV3.GuideRateDeclination
        //     property. These two rates may be tied together into a single rate, depending
        //     on the driver's implementation and the capabilities of the telescope.
        public void PulseGuide(GuideDirections Direction, int Duration)
        {
            tl.LogMessage("PulseGuide", "Direction => " + Direction);
            tl.LogMessage("PulseGuide", "Duration => " + Duration + " msec.");
            mTxPkt[1] = (byte)mSeq++;
            mTxPkt[2] = (byte)Direction;
            byte[] duration = BitConverter.GetBytes(Duration);
            for (int i = 0; i < 4; ++i)
            {
                mTxPkt[3 + i] = duration[i];
            }
            if (Direction == GuideDirections.guideEast || Direction == GuideDirections.guideWest)
            {
                mTxPkt[7] = (byte)(mGuideRateRA * 10.0);
            }
            else
            {
                mTxPkt[7] = (byte)(mGuideRateDEC * 10.0);
            }
            if (mDevice != null)
            {
                for (int i = 0; i < 3; ++i)
                {
                    mSender.Send(mTxPkt, mTxPkt.Length, mDevice);
                }
            }
            else
            {
                mSender.Send(mTxPkt, mTxPkt.Length, new IPEndPoint(IPAddress.Broadcast, TX_PORT));
                tl.LogMessage("PulseGuide", "WARNING: Opensynscan not discovered");
            }
            mLastCommandEndTime = DateTime.Now.AddMilliseconds(Duration);           
        }


        //
        // Resumen:
        //     True if this telescope is capable of software-pulsed guiding (via the ASCOM.DeviceInterface.ITelescopeV3.PulseGuide(ASCOM.DeviceInterface.GuideDirections,System.Int32)
        //     method)
        //
        // Comentarios:
        //     Must be implemented, must not throw a PropertyNotImplementedException.  May
        //     raise an error if the telescope is not connected.
        public bool CanPulseGuide
        {
            get
            {
                tl.LogMessage("CanPulseGuide", "Get - " + true.ToString());
                return true;
            }
        }

        //
        // Resumen:
        //     True if the guide rate properties used for ASCOM.DeviceInterface.ITelescopeV3.PulseGuide(ASCOM.DeviceInterface.GuideDirections,System.Int32)
        //     can ba adjusted.
        //
        // Comentarios:
        //     Must be implemented, must not throw a PropertyNotImplementedException.  May
        //     raise an error if the telescope is not connected.
        //     This is only available for telescope InterfaceVersions 2 and 3
        public bool CanSetGuideRates
        {
            get
            {
                tl.LogMessage("CanSetGuideRates", "Get - " + true.ToString());
                return true;
            }
        }

        //
        // Resumen:
        //     The current Declination movement rate offset for telescope guiding (degrees/sec)
        //
        // Excepciones:
        //   ASCOM.PropertyNotImplementedException:
        //     If the property is not implemented
        //
        //   ASCOM.InvalidValueException:
        //     If an invalid guide rate is set.
        //
        // Comentarios:
        //     This is the rate for both hardware/relay guiding and the PulseGuide() method.
        //     This is only available for telescope InterfaceVersions 2 and 3
        //     NOTES: To discover whether this feature is supported, test the ASCOM.DeviceInterface.ITelescopeV3.CanSetGuideRates
        //     property. The supported range of this property is telescope specific, however,
        //     if this feature is supported, it can be expected that the range is sufficient
        //     to allow correction of guiding errors caused by moderate misalignment and
        //     periodic error. If a telescope does not support separate guiding rates in
        //     Right Ascension and Declination, then it is permissible for ASCOM.DeviceInterface.ITelescopeV3.GuideRateRightAscension
        //     and GuideRateDeclination to be tied together.  In this case, changing one
        //     of the two properties will cause a change in the other. Mounts must start
        //     up with a known or default declination guide rate, and this property must
        //     return that known/default guide rate until changed.
        double mGuideRateDEC = 0.7;
        public double GuideRateDeclination
        {
            get
            {
                tl.LogMessage("GuideRateDEC Get - ", mGuideRateDEC.ToString());
                return mGuideRateDEC;
            }
            set
            {
                tl.LogMessage("GuideRateDEC Set - ", mGuideRateDEC.ToString() + " => " + value);
                //mGuideRateDEC = value;
            }
        }

        //
        // Resumen:
        //     The current Right Ascension movement rate offset for telescope guiding (degrees/sec)
        //
        // Excepciones:
        //   ASCOM.PropertyNotImplementedException:
        //     If the property is not implemented
        //
        //   ASCOM.InvalidValueException:
        //     If an invalid guide rate is set.
        //
        // Comentarios:
        //     This is the rate for both hardware/relay guiding and the PulseGuide() method.
        //     This is only available for telescope InterfaceVersions 2 and 3
        //     NOTES: To discover whether this feature is supported, test the ASCOM.DeviceInterface.ITelescopeV3.CanSetGuideRates
        //     property. The supported range of this property is telescope specific, however,
        //     if this feature is supported, it can be expected that the range is sufficient
        //     to allow correction of guiding errors caused by moderate misalignment and
        //     periodic error. If a telescope does not support separate guiding rates in
        //     Right Ascension and Declination, then it is permissible for GuideRateRightAscension
        //     and ASCOM.DeviceInterface.ITelescopeV3.GuideRateDeclination to be tied together.
        //     In this case, changing one of the two properties will cause a change in the
        //     other. Mounts must start up with a known or default right ascension guide
        //     rate, and this property must return that known/default guide rate until changed.
        double mGuideRateRA = 0.7;
        public double GuideRateRightAscension
        {
            get
            {
                tl.LogMessage("GuideRateRightRA Get - ", mGuideRateRA.ToString());
                return mGuideRateRA;
            }
            set
            {
                tl.LogMessage("GuideRateRightRA Set - ", mGuideRateRA.ToString() + " => " + value);
                //mGuideRateRA = value;
            }
        }

        //
        // Resumen:
        //     True if a ASCOM.DeviceInterface.ITelescopeV3.PulseGuide(ASCOM.DeviceInterface.GuideDirections,System.Int32)
        //     command is in progress, False otherwise
        //
        // Excepciones:
        //   ASCOM.PropertyNotImplementedException:
        //     If ASCOM.DeviceInterface.ITelescopeV3.CanPulseGuide is False
        //
        // Comentarios:
        //     Raises an error if the value of the ASCOM.DeviceInterface.ITelescopeV3.CanPulseGuide
        //     property is false (the driver does not support the ASCOM.DeviceInterface.ITelescopeV3.PulseGuide(ASCOM.DeviceInterface.GuideDirections,System.Int32)
        //     method).
        public bool IsPulseGuiding
        {
            get
            {
                bool is_guiding = (DateTime.Now < mLastCommandEndTime);
                tl.LogMessage("IsPulseGuiding Get - ", is_guiding.ToString());
                return is_guiding;
            }
        }

        //////////////////////////////////////////////
       
        public void AbortSlew()
        {
            tl.LogMessage("AbortSlew", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("AbortSlew");
        }

        public AlignmentModes AlignmentMode
        {
            get
            {
                tl.LogMessage("AlignmentMode Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("AlignmentMode", false);
            }
        }

        public double Altitude
        {
            get
            {
                tl.LogMessage("Altitude", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Altitude", false);
            }
        }

        public double ApertureArea
        {
            get
            {
                tl.LogMessage("ApertureArea Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("ApertureArea", false);
            }
        }

        public double ApertureDiameter
        {
            get
            {
                tl.LogMessage("ApertureDiameter Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("ApertureDiameter", false);
            }
        }

        public bool AtHome
        {
            get
            {
                tl.LogMessage("AtHome", "Get - " + false.ToString());
                return false;
            }
        }

        public bool AtPark
        {
            get
            {
                tl.LogMessage("AtPark", "Get - " + false.ToString());
                return false;
            }
        }

        public IAxisRates AxisRates(TelescopeAxes Axis)
        {
            tl.LogMessage("AxisRates", "Get - " + Axis.ToString());
            return new AxisRates(Axis);
        }

        public double Azimuth
        {
            get
            {
                tl.LogMessage("Azimuth Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Azimuth", false);
            }
        }

        public bool CanFindHome
        {
            get
            {
                tl.LogMessage("CanFindHome", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanMoveAxis(TelescopeAxes Axis)
        {
            tl.LogMessage("CanMoveAxis", "Get - " + Axis.ToString());
            switch (Axis)
            {
                case TelescopeAxes.axisPrimary: return false;
                case TelescopeAxes.axisSecondary: return false;
                case TelescopeAxes.axisTertiary: return false;
                default: throw new InvalidValueException("CanMoveAxis", Axis.ToString(), "0 to 2");
            }
        }

        public bool CanPark
        {
            get
            {
                tl.LogMessage("CanPark", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetDeclinationRate
        {
            get
            {
                tl.LogMessage("CanSetDeclinationRate", "Get - " + false.ToString());
                return false;
            }
        }


        public bool CanSetPark
        {
            get
            {
                tl.LogMessage("CanSetPark", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetPierSide
        {
            get
            {
                tl.LogMessage("CanSetPierSide", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetRightAscensionRate
        {
            get
            {
                tl.LogMessage("CanSetRightAscensionRate", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetTracking
        {
            get
            {
                tl.LogMessage("CanSetTracking", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSlew
        {
            get
            {
                tl.LogMessage("CanSlew", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSlewAltAz
        {
            get
            {
                tl.LogMessage("CanSlewAltAz", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSlewAltAzAsync
        {
            get
            {
                tl.LogMessage("CanSlewAltAzAsync", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSlewAsync
        {
            get
            {
                tl.LogMessage("CanSlewAsync", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSync
        {
            get
            {
                tl.LogMessage("CanSync", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSyncAltAz
        {
            get
            {
                tl.LogMessage("CanSyncAltAz", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanUnpark
        {
            get
            {
                tl.LogMessage("CanUnpark", "Get - " + false.ToString());
                return false;
            }
        }

        public double Declination
        {
            get
            {
                double declination = 0.0;
                tl.LogMessage("Declination", "Get - " + utilities.DegreesToDMS(declination, ":", ":"));
                return declination;
            }
        }

        public double DeclinationRate
        {
            get
            {
                double declination = 0.0;
                tl.LogMessage("DeclinationRate", "Get - " + declination.ToString());
                return declination;
            }
            set
            {
                tl.LogMessage("DeclinationRate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DeclinationRate", true);
            }
        }

        public PierSide DestinationSideOfPier(double RightAscension, double Declination)
        {
            tl.LogMessage("DestinationSideOfPier Get", "Not implemented");
            throw new ASCOM.PropertyNotImplementedException("DestinationSideOfPier", false);
        }

        public bool DoesRefraction
        {
            get
            {
                tl.LogMessage("DoesRefraction Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DoesRefraction", false);
            }
            set
            {
                tl.LogMessage("DoesRefraction Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DoesRefraction", true);
            }
        }

        public EquatorialCoordinateType EquatorialSystem
        {
            get
            {
                EquatorialCoordinateType equatorialSystem = EquatorialCoordinateType.equLocalTopocentric;
                tl.LogMessage("DeclinationRate", "Get - " + equatorialSystem.ToString());
                return equatorialSystem;
            }
        }

        public void FindHome()
        {
            tl.LogMessage("FindHome", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("FindHome");
        }

        public double FocalLength
        {
            get
            {
                tl.LogMessage("FocalLength Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("FocalLength", false);
            }
        }

       

        public void MoveAxis(TelescopeAxes Axis, double Rate)
        {
            tl.LogMessage("MoveAxis", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("MoveAxis");
        }

        public void Park()
        {
            tl.LogMessage("Park", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("Park");
        }



        public double RightAscension
        {
            get
            {
                double rightAscension = 0.0;
                tl.LogMessage("RightAscension", "Get - " + utilities.HoursToHMS(rightAscension));
                return rightAscension;
            }
        }

        public double RightAscensionRate
        {
            get
            {
                double rightAscensionRate = 0.0;
                tl.LogMessage("RightAscensionRate", "Get - " + rightAscensionRate.ToString());
                return rightAscensionRate;
            }
            set
            {
                tl.LogMessage("RightAscensionRate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("RightAscensionRate", true);
            }
        }

        public void SetPark()
        {
            tl.LogMessage("SetPark", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SetPark");
        }

        public PierSide SideOfPier
        {
            get
            {
                tl.LogMessage("SideOfPier Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SideOfPier", false);
            }
            set
            {
                tl.LogMessage("SideOfPier Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SideOfPier", true);
            }
        }

        public double SiderealTime
        {
            get
            {
                // Now using NOVAS 3.1
                double siderealTime = 0.0;
                using (var novas = new ASCOM.Astrometry.NOVAS.NOVAS31())
                {
                    var jd = utilities.DateUTCToJulian(DateTime.UtcNow);
                    novas.SiderealTime(jd, 0, novas.DeltaT(jd),
                        ASCOM.Astrometry.GstType.GreenwichApparentSiderealTime,
                        ASCOM.Astrometry.Method.EquinoxBased,
                        ASCOM.Astrometry.Accuracy.Reduced, ref siderealTime);
                }

                // Allow for the longitude
                siderealTime += SiteLongitude / 360.0 * 24.0;

                // Reduce to the range 0 to 24 hours
                siderealTime = astroUtilities.ConditionRA(siderealTime);

                tl.LogMessage("SiderealTime", "Get - " + siderealTime.ToString());
                return siderealTime;
            }
        }

        public double SiteElevation
        {
            get
            {
                tl.LogMessage("SiteElevation Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteElevation", false);
            }
            set
            {
                tl.LogMessage("SiteElevation Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteElevation", true);
            }
        }

        public double SiteLatitude
        {
            get
            {
                tl.LogMessage("SiteLatitude Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteLatitude", false);
            }
            set
            {
                tl.LogMessage("SiteLatitude Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteLatitude", true);
            }
        }

        public double SiteLongitude
        {
            get
            {
                tl.LogMessage("SiteLongitude Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteLongitude", false);
            }
            set
            {
                tl.LogMessage("SiteLongitude Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteLongitude", true);
            }
        }

        public short SlewSettleTime
        {
            get
            {
                tl.LogMessage("SlewSettleTime Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SlewSettleTime", false);
            }
            set
            {
                tl.LogMessage("SlewSettleTime Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SlewSettleTime", true);
            }
        }

        public void SlewToAltAz(double Azimuth, double Altitude)
        {
            tl.LogMessage("SlewToAltAz", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAltAz");
        }

        public void SlewToAltAzAsync(double Azimuth, double Altitude)
        {
            tl.LogMessage("SlewToAltAzAsync", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAltAzAsync");
        }

        public void SlewToCoordinates(double RightAscension, double Declination)
        {
            tl.LogMessage("SlewToCoordinates", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToCoordinates");
        }

        public void SlewToCoordinatesAsync(double RightAscension, double Declination)
        {
            tl.LogMessage("SlewToCoordinatesAsync", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToCoordinatesAsync");
        }

        public void SlewToTarget()
        {
            tl.LogMessage("SlewToTarget", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToTarget");
        }

        public void SlewToTargetAsync()
        {
            tl.LogMessage("SlewToTargetAsync", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToTargetAsync");
        }

        public bool Slewing
        {
            get
            {
                tl.LogMessage("Slewing Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Slewing", false);
            }
        }

        public void SyncToAltAz(double Azimuth, double Altitude)
        {
            tl.LogMessage("SyncToAltAz", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SyncToAltAz");
        }

        public void SyncToCoordinates(double RightAscension, double Declination)
        {
            tl.LogMessage("SyncToCoordinates", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SyncToCoordinates");
        }

        public void SyncToTarget()
        {
            tl.LogMessage("SyncToTarget", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SyncToTarget");
        }

        public double TargetDeclination
        {
            get
            {
                tl.LogMessage("TargetDeclination Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TargetDeclination", false);
            }
            set
            {
                tl.LogMessage("TargetDeclination Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TargetDeclination", true);
            }
        }

        public double TargetRightAscension
        {
            get
            {
                tl.LogMessage("TargetRightAscension Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TargetRightAscension", false);
            }
            set
            {
                tl.LogMessage("TargetRightAscension Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TargetRightAscension", true);
            }
        }

        public bool Tracking
        {
            get
            {
                bool tracking = true;
                tl.LogMessage("Tracking", "Get - " + tracking.ToString());
                return tracking;
            }
            set
            {
                tl.LogMessage("Tracking Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Tracking", true);
            }
        }

        public DriveRates TrackingRate
        {
            get
            {
                tl.LogMessage("TrackingRate Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TrackingRate", false);
            }
            set
            {
                tl.LogMessage("TrackingRate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TrackingRate", true);
            }
        }

        public ITrackingRates TrackingRates
        {
            get
            {
                ITrackingRates trackingRates = new TrackingRates();
                tl.LogMessage("TrackingRates", "Get - ");
                foreach (DriveRates driveRate in trackingRates)
                {
                    tl.LogMessage("TrackingRates", "Get - " + driveRate.ToString());
                }
                return trackingRates;
            }
        }

        public DateTime UTCDate
        {
            get
            {
                DateTime utcDate = DateTime.UtcNow;
                tl.LogMessage("TrackingRates", "Get - " + String.Format("MM/dd/yy HH:mm:ss", utcDate));
                return utcDate;
            }
            set
            {
                tl.LogMessage("UTCDate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("UTCDate", true);
            }
        }

        public void Unpark()
        {
            tl.LogMessage("Unpark", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("Unpark");
        }

        #endregion

        #region Private properties and methods
        // here are some useful properties and methods that can be used as required
        // to help with driver development

        #region ASCOM Registration

        // Register or unregister driver for ASCOM. This is harmless if already
        // registered or unregistered. 
        //
        /// <summary>
        /// Register or unregister the driver with the ASCOM Platform.
        /// This is harmless if the driver is already registered/unregistered.
        /// </summary>
        /// <param name="bRegister">If <c>true</c>, registers the driver, otherwise unregisters it.</param>
        private static void RegUnregASCOM(bool bRegister)
        {
            using (var P = new ASCOM.Utilities.Profile())
            {
                P.DeviceType = "Telescope";
                if (bRegister)
                {
                    P.Register(driverID, driverDescription);
                }
                else
                {
                    P.Unregister(driverID);
                }
            }
        }

        /// <summary>
        /// This function registers the driver with the ASCOM Chooser and
        /// is called automatically whenever this class is registered for COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is successfully built.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During setup, when the installer registers the assembly for COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually register a driver with ASCOM.
        /// </remarks>
        [ComRegisterFunction]
        public static void RegisterASCOM(Type t)
        {
            RegUnregASCOM(true);
        }

        /// <summary>
        /// This function unregisters the driver from the ASCOM Chooser and
        /// is called automatically whenever this class is unregistered from COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is cleaned or prior to rebuilding.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During uninstall, when the installer unregisters the assembly from COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually unregister a driver from ASCOM.
        /// </remarks>
        [ComUnregisterFunction]
        public static void UnregisterASCOM(Type t)
        {
            RegUnregASCOM(false);
        }

        #endregion

        /// <summary>
        /// Returns true if there is a valid connection to the driver hardware
        /// </summary>
        private bool IsConnected
        {
            get
            {
                // TODO check that the driver hardware connection exists and is connected to the hardware
                return connectedState;
            }
        }

        /// <summary>
        /// Use this function to throw an exception if we aren't connected to the hardware
        /// </summary>
        /// <param name="message"></param>
        private void CheckConnected(string message)
        {
            if (!IsConnected)
            {
                throw new ASCOM.NotConnectedException(message);
            }
        }

        /// <summary>
        /// Read the device configuration from the ASCOM Profile store
        /// </summary>
        internal void ReadProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Telescope";
                tl.Enabled = true;
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal void WriteProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Telescope";
            }
        }

        /// <summary>
        /// Log helper function that takes formatted strings and arguments
        /// </summary>
        /// <param name="identifier"></param>
        /// <param name="message"></param>
        /// <param name="args"></param>
        internal static void LogMessage(string identifier, string message, params object[] args)
        {
            var msg = string.Format(message, args);
            tl.LogMessage(identifier, msg);
        }
        #endregion
    }
}
