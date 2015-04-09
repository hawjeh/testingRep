using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ININ.Audio;
using ININ.InteractionClient.Voicemail.Properties;
using ININ.Windows.Forms;
using ININ.Windows.Forms.Audio;

namespace ININ.InteractionClient.Voicemail
{
    /// <summary>
    /// Summary description for VoicemailPlayerControl.
    /// </summary>
    public class VoicemailPlayerControl : UserControl
    {
        // Used for checking if playback was disconnected
        private const int E_I3ERR_TS_CANCELLED = unchecked((int)0x80040011);

        private const int REWIND_FASTFORWARD_SECONDS = 3;

        private VoicemailAttachmentContainer _container;
        private IAudio _audio;
        private bool _timerStarted;
        private bool _playAutomatically;

        private int _initialHeight;

        private readonly IUserInterfaceContext _userInterfaceContext;
        private PlayControl _playControl;
        private PictureButton _deleteButton;
        private ToolTip _toolTip;
        private Panel _locationPanel;
        private PlaybackLocation _playbackLocation;
        private PlaybackInfo _playbackInfo;
        private Timer _updateTimer;
        private IContainer components;
        private DownloadManager _downloadManager;

        [Obsolete]
        public VoicemailPlayerControl()
        {
            // Explicitly for the Windows.Forms Form Designer.
            InitializeComponent();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VoicemailPlayerControl"/> class. 
        /// </summary>
        public VoicemailPlayerControl(IUserInterfaceContext userInterfaceContext)
        {
            _userInterfaceContext = userInterfaceContext;
            ////_userInterfaceContext = ServiceLocator.Current.GetInstance<IUserInterfaceContext>();

            // This call is required by the Windows.Forms Form Designer.
            InitializeComponent();
            
            _playbackInfo.RefreshList += OnPlaybackInfoRefreshList;
            _playbackInfo.SelectedPlaybackDeviceChanged += OnPlaybackDeviceChanged;

            _playbackLocation.ValueChanged += OnPlaybackLocationChanged;
            _playbackLocation.MouseDown += OnPlaybackLocationMouseDown;
            _playbackLocation.MouseUp += OnPlaybackLocationMouseUp;

            _playControl.Play += OnPlayControlPlay;
            _playControl.Stop += OnPlayControlStop;
            _playControl.Rewind += OnPlayControlRewind;
            _playControl.FastForward += OnPlayControlFastForward;
            _playControl.VolumeChanged += OnPlayControlVolumeChanged;
            _playControl.SpeedChanged += OnPlayControlSpeedChanged;

            _playControl.AddControl(_deleteButton);
            _toolTip.SetToolTip(_deleteButton, Resources.DELETE_TOOLTIP);

            OnContainerChanged();
        }

        public DownloadManager DownloadManager
        {
            get { return _downloadManager; }
            set
            {
                if (value != null)
                {
                    _downloadManager = value;
                    UnHookDownloadManager(_downloadManager);
                    HookUpDownloadManager(_downloadManager);
                }
            }
        }

        private void HookUpDownloadManager(DownloadManager downloadManager)
        {
            if (downloadManager == null) return;

            downloadManager.RemovingLocalFile += OnRemovingLocalFile;
            downloadManager.DownloadCompleted += OnDownloadCompleted;
        }

        private void UnHookDownloadManager(DownloadManager downloadManager)
        {
            if (downloadManager == null) return;

            downloadManager.RemovingLocalFile -= OnRemovingLocalFile;
            downloadManager.DownloadCompleted -= OnDownloadCompleted;
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }

                UnHookDownloadManager(_downloadManager);
                CleanUpAudio();
            }

            base.Dispose(disposing);
        }

        #region Component Designer generated code
        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(VoicemailPlayerControl));
            this._playControl = new ININ.Windows.Forms.Audio.PlayControl();
            this._deleteButton = new ININ.Windows.Forms.PictureButton();
            this._locationPanel = new System.Windows.Forms.Panel();
            this._playbackLocation = new ININ.Windows.Forms.Audio.PlaybackLocation();
            this._playbackInfo = new ININ.InteractionClient.Voicemail.PlaybackInfo();
            this._updateTimer = new System.Windows.Forms.Timer(this.components);
            this._toolTip = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this._deleteButton)).BeginInit();
            this._locationPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // playControl
            // 
            this._playControl.Dock = System.Windows.Forms.DockStyle.Bottom;
            this._playControl.Location = new System.Drawing.Point(0, 40);
            this._playControl.Mute = false;
            this._playControl.Name = "_playControl";
            this._playControl.PlayControlMode = ININ.Windows.Forms.Audio.PlayControlMode.Full;
            this._playControl.Size = new System.Drawing.Size(416, 38);
            this._playControl.Speed = ININ.Windows.Forms.Audio.SpeedLevel.Normal;
            this._playControl.TabIndex = 1;
            this._playControl.Volume = 100;
            // 
            // deleteButton
            // 
            this._deleteButton.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this._deleteButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this._deleteButton.DefaultImage = ((System.Drawing.Image)(resources.GetObject("_deleteButton.DefaultImage")));
            this._deleteButton.DisabledImage = ((System.Drawing.Image)(resources.GetObject("_deleteButton.DisabledImage")));
            this._deleteButton.HoverImage = ((System.Drawing.Image)(resources.GetObject("_deleteButton.HoverImage")));
            this._deleteButton.Image = ((System.Drawing.Image)(resources.GetObject("_deleteButton.Image")));
            this._deleteButton.Location = new System.Drawing.Point(320, 0);
            this._deleteButton.Margin = new System.Windows.Forms.Padding(46, 3, 3, 3);
            this._deleteButton.Name = "_deleteButton";
            this._deleteButton.Size = new System.Drawing.Size(24, 24);
            this._deleteButton.TabIndex = 1;
            this._deleteButton.TabStop = false;
            this._deleteButton.Click += new System.EventHandler(this.DeleteButtonClick);
            // 
            // locationPanel
            // 
            this._locationPanel.Controls.Add(this._playbackLocation);
            this._locationPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this._locationPanel.Location = new System.Drawing.Point(0, 30);
            this._locationPanel.Name = "_locationPanel";
            this._locationPanel.Size = new System.Drawing.Size(416, 10);
            this._locationPanel.TabIndex = 6;
            this._locationPanel.VisibleChanged += new System.EventHandler(this.OnLocationPanelVisibleChanged);
            // 
            // playbackLocation
            // 
            this._playbackLocation.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
            this._playbackLocation.BorderColor = System.Drawing.SystemColors.ControlDarkDark;
            this._playbackLocation.Cursor = System.Windows.Forms.Cursors.Hand;
            this._playbackLocation.EndColor = System.Drawing.SystemColors.ControlDark;
            this._playbackLocation.GradientDirection = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
            this._playbackLocation.Location = new System.Drawing.Point(2, 0);
            this._playbackLocation.Maximum = ((long)(100));
            this._playbackLocation.Minimum = ((long)(0));
            this._playbackLocation.Name = "_playbackLocation";
            this._playbackLocation.ShowBorder = true;
            this._playbackLocation.Size = new System.Drawing.Size(412, 10);
            this._playbackLocation.StartColor = System.Drawing.SystemColors.Control;
            this._playbackLocation.TabIndex = 0;
            this._playbackLocation.Value = ((long)(0));
            // 
            // playbackInfo
            // 
            this._playbackInfo.Dock = System.Windows.Forms.DockStyle.Bottom;
            this._playbackInfo.SpecialMessage = Resources.VOICEMAIL_DOWNLOAD_REQUIRED;
            this._playbackInfo.ShowSpecialMessage = false;
            this._playbackInfo.Location = new System.Drawing.Point(0, 2);
            this._playbackInfo.Name = "_playbackInfo";
            this._playbackInfo.Size = new System.Drawing.Size(416, 28);
            this._playbackInfo.TabIndex = 0;
            // 
            // updateTimer
            // 
            this._updateTimer.Tick += new System.EventHandler(this.UpdateTimer_Tick);
            // 
            // VoicemailPlayerControl
            // 
            this.Controls.Add(this._playbackInfo);
            this.Controls.Add(this._locationPanel);
            this.Controls.Add(this._playControl);
            this.Name = "VoicemailPlayerControl";
            this.Size = new System.Drawing.Size(416, 78);
            this.Load += new System.EventHandler(this.VMPlayer_Load);
            ((System.ComponentModel.ISupportInitialize)(this._deleteButton)).EndInit();
            this._locationPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        #endregion

        #region Events

        /// <summary>
        /// Occurs when the user chooses to refresh the voicemail list.
        /// </summary>
        public event EventHandler RefreshList;

        private void OnRefreshList()
        {
            try
            {
                if (RefreshList != null)
                {
                    RefreshList(this, EventArgs.Empty);
                }
            }
            catch (Exception ex)
            {
                TraceTopics.Voicemail.exception(ex);
            }
        }

        /// <summary>
        /// A DeleteItem event occurs when a voicemail has been deleted from the player
        /// </summary>
        public event EventHandler DeleteItem;

        private void DeleteButtonClick(object sender, EventArgs e)
        {
            try
            {
                EventHandler deleteEventHandler = DeleteItem;
                if (deleteEventHandler != null)
                {
                    deleteEventHandler(this, EventArgs.Empty);
                }
            }
            catch (Exception ex)
            {
                TraceTopics.Voicemail.exception(ex, "Problem trying to delete voicemail.");
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the current voicemail attachment to play.
        /// </summary>
        public VoicemailAttachmentContainer VoicemailAttachmentContainer
        {
            get { return _container; }
            set
            {
                if (_container != null)
                {
                    if (!_container.Equals(value))
                    {
                        OnContainerChanging();
                        _container = value;
                        OnContainerChanged();
                    }
                }
                else if (value != null)
                {
                    // _container is null, value is not, assign
                    OnContainerChanging();
                    _container = value;
                    OnContainerChanged();
                }
            }
        }

        public bool DeleteEnabled
        {
            get
            {
                return _deleteButton.Enabled;
            }
            set
            {
                _deleteButton.Enabled = value;
            }
        }

        #endregion

        private void ClearAudio()
        {
            if (_audio != null)
            {
                CleanUpAudio();
                OnAudioChanged();
            }
        }

        private void CleanUpAudio()
        {
            if (_audio != null)
            {
                StopAudioAndResetControls();

                _audio.Pausing -= OnAudioPausing;
                _audio.Starting -= OnAudioStarting;
                _audio.Stopping -= OnAudioStopping;

                var disposable = _audio as IDisposable;
                if (disposable != null)
                {
                    disposable.Dispose();
                }

                _audio = null;
            }
        }

        private void OnAudioChanged()
        {
            if (_audio != null)
            {
                // Reset volume and playback speed of audio based on controls
                _audio.Volume = _playControl.Volume;
                switch (_playControl.Speed)
                {
                    case SpeedLevel.Half:
                        _audio.Frequency = -50;
                        break;
                    case SpeedLevel.Normal:
                        _audio.Frequency = 0;
                        break;
                    case SpeedLevel.Faster:
                        _audio.Frequency = 50;
                        break;
                    case SpeedLevel.Fastest:
                        _audio.Frequency = 100;
                        break;
                }
            }
        }

        private void SetupAudio()
        {
            // Create audio based on local file
            try
            {
                if (_container != null)
                {
                    string file = _downloadManager[_container];
                    if (file.Length > 0)
                    {
                        _audio = AudioFactory.CreateAudio(file);

                        if (_audio != null)
                        {
                            _audio.Pausing += OnAudioPausing;
                            _audio.Starting += OnAudioStarting;
                            _audio.Stopping += OnAudioStopping;
                        }

                        OnAudioChanged();
                    }
                }
            }
            catch (Exception ex)
            {
                // Trap all exceptions, set audio to null
                TraceTopics.Voicemail.exception(ex);
                _audio = null;
            }
        }

        /// <summary>
        /// This is the opportunity to do any modifications to the existing container
        /// </summary>
        private void OnContainerChanging()
        {
            if (_container != null)
            {
                // Stop audio playback
                switch (_playbackInfo.SelectedPlaybackDevice)
                {
                    case PlaybackDevice.RemoteTelephone:
                        _container.Message.StopPlayToNumberAsync(StopPlayToNumberCompleted, null);
                        break;
                    case PlaybackDevice.TelephoneHandset:
                        _container.Message.StopPlayToStationAsync(StopPlayToStationCompleted, null);
                        break;
                }

                _container.Message.PlaybackComplete -= OnMessagePlaybackComplete;
            }
        }

        private void OnContainerChanged()
        {
            // Stop all current playback immediately, then evaluate the new container.
            ClearAudio();

            SetupAudio();

            bool exists = _container != null;
            bool audioExists = _audio != null;

            if (exists)
            {
                _container.Message.PlaybackComplete += OnMessagePlaybackComplete;
            }

            // Reset _playAutomatically
            _playAutomatically = false;

            _playbackInfo.SpecialMessage = Resources.VOICEMAIL_DOWNLOAD_REQUIRED;
            _playbackInfo.ShowSpecialMessage = exists && !audioExists;

            // Adjust play controls based on selected playback device
            AdjustPlayControls();

            _playbackInfo.Enabled = true;
        }

        private void AdjustPlayControls()
        {
            // Enable play control and playback location
            if (_playControl.Enabled != (_container != null))
            {
                _playControl.Enabled = _container != null;
                _deleteButton.Enabled = _container != null;
            }

            if (_playbackLocation.Enabled != (_audio != null))
            {
                _playbackLocation.Enabled = _audio != null;
            }

            _playbackLocation.Value = 0;

            switch (_playbackInfo.SelectedPlaybackDevice)
            {
                case PlaybackDevice.PCSpeakers:
                    // Show all buttons, location slider, and enable/disable buttons appropriately
                    _locationPanel.Visible = true;
                    _playControl.PlayControlMode = PlayControlMode.Full;
                    _playControl.SetButtonEnabled(PlaybackButton.All & ~PlaybackButton.Play, (_audio != null) && (_container != null));

                    // Enable the play button if DirectX 9 or higher is available.
                    // If DX9 isn't available, don't even allow the user to download the voicemail (what's the point?)
                    _playControl.SetButtonEnabled(PlaybackButton.Play, RegistryUtils.IsDirectX9OrHigher());
                    break;
                case PlaybackDevice.RemoteTelephone:
                case PlaybackDevice.TelephoneHandset:
                    // Show play and stop buttons only, and enable/disable buttons appropriately, hide location slider
                    _locationPanel.Visible = false;
                    _playControl.PlayControlMode = PlayControlMode.PlayStop;
                    if (_container != null)
                    {
                        _playControl.SetButtonEnabled(PlaybackButton.Play, true);
                        _playControl.SetButtonEnabled(PlaybackButton.Stop, false);
                    }
                    else
                    {
                        _playControl.SetButtonEnabled(PlaybackButton.Play | PlaybackButton.Stop, false);
                    }

                    break;
            }
        }

        private void OnPlaybackInfoRefreshList(object sender, EventArgs e)
        {
            // Bubble
            OnRefreshList();
        }

        private void OnPlaybackDeviceChanged(object sender, EventArgs e)
        {
            AdjustPlayControls();

            if ((_playbackInfo.SelectedPlaybackDevice != PlaybackDevice.PCSpeakers) && (_audio != null))
            {
                // Stop the audio playback
                StopAudioAndResetControls();
            }
        }

        private void OnPlayControlPlay(object sender, EventArgs e)
        {
            if (_container != null)
            {
                try
                {
                    switch (_playbackInfo.SelectedPlaybackDevice)
                    {
                        case PlaybackDevice.PCSpeakers:
                            if (_audio != null)
                            {
                                _audio.Play();
                            }
                            else
                            {
                                // Initiate download, begin playing when downloaded
                                _playAutomatically = true;
                                string file = _downloadManager[_container];
                                if (file.Length == 0)
                                {
                                    if (!_downloadManager.IsEnqueued(_container))
                                    {
                                        _downloadManager.Enqueue(_container);
                                    }

                                    _playbackInfo.SpecialMessage = Resources.VOICEMAIL_DOWNLOADING;
                                }
                            }

                            break;
                        case PlaybackDevice.RemoteTelephone:
                            if (_playbackInfo.Target.Length > 0)
                            {
                                _container.Message.PlayToNumberAsync(_container.Attachment, _playbackInfo.Target, true, true, PlayToNumberCompleted, null);
                                _playControl.SetButtonEnabled(PlaybackButton.Play, false);
                                _playControl.SetButtonEnabled(PlaybackButton.Stop, true);

                                _playbackInfo.Enabled = false;
                            }

                            break;
                        case PlaybackDevice.TelephoneHandset:
                            if (_playbackInfo.StationName.Length > 0)
                            {
                                _container.Message.PlayToStationAsync(_container.Attachment, _playbackInfo.StationName, true, true, PlayToStationCompleted, null);
                                _playControl.SetButtonEnabled(PlaybackButton.Play, false);
                                _playControl.SetButtonEnabled(PlaybackButton.Stop, true);

                                _playbackInfo.Enabled = false;
                            }

                            break;
                    }
                }
                catch (Exception ex)
                {
                    _playbackInfo.PlaybackError = Resources.VOICEMAIL_PCSPEAKER_ERROR;
                    _playbackInfo.ShowPlaybackError = true;
                    TraceTopics.Voicemail.exception(ex);
                }
            }
        }

        private void PlayToNumberCompleted(object sender, AsyncCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                TraceTopics.Voicemail.warning("Error received from VoicemailMessage.PlayToNumberCompleted, reverting UI and continuing");
                EnablePlayAfterError(e.Error);
            }
        }

        private void PlayToStationCompleted(object sender, AsyncCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                TraceTopics.Voicemail.warning("Error received from VoicemailMessage.PlayToStationCompleted, reverting UI and continuing");
                EnablePlayAfterError(e.Error);
            }
        }

        private void EnablePlayAfterError(Exception exception)
        {
            if (IsHandleCreated)
            {
                if (InvokeRequired)
                {
                    MethodInvoker mi = () => EnablePlayAfterError(exception);
                    BeginInvoke(mi);
                }
                else
                {
                    // Reset the button state
                    _playControl.SetButtonEnabled(PlaybackButton.Play, true);
                    _playControl.SetButtonEnabled(PlaybackButton.Stop, false);

                    _playbackInfo.Enabled = true;

                    if ((exception != null) && (Marshal.GetHRForException(exception) == E_I3ERR_TS_CANCELLED))
                    {
                        TraceTopics.Voicemail.note("TS is reporting a canceled error. This probably means that the user clicked disconnect on the outbound call to play the voicemail. We'll suppress this message from the user since it was probably intentional.");
                    }
                    else
                    {
                        LocalizedMessageBox.Show(String.Format("{0}{1}{2}", Resources.VOICEMAIL_PLAYBACK_ERROR, Environment.NewLine, Resources.VOICEMAIL_PLAYBACK_ERROR_SOLUTION), Resources.VOICEMAIL_PLAYER_ERROR_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private void OnPlayControlStop(object sender, EventArgs e)
        {
            switch (_playbackInfo.SelectedPlaybackDevice)
            {
                case PlaybackDevice.PCSpeakers:
                    if (_audio != null)
                    {
                        StopAudioAndResetControls();
                    }
                    else
                    {
                        _playControl.SetButtonEnabled(PlaybackButton.All, false);
                    }

                    break;
                case PlaybackDevice.RemoteTelephone:
                    if (_container != null)
                    {
                        _container.Message.StopPlayToNumberAsync(StopPlayToNumberCompleted, null);
                    }

                    _playControl.SetButtonEnabled(PlaybackButton.Play, true);
                    _playControl.SetButtonEnabled(PlaybackButton.Stop, false);
                    _playbackInfo.Enabled = true;
                    break;
                case PlaybackDevice.TelephoneHandset:
                    if (_container != null)
                    {
                        _container.Message.StopPlayToStationAsync(StopPlayToStationCompleted, null);
                    }

                    _playControl.SetButtonEnabled(PlaybackButton.Play, true);
                    _playControl.SetButtonEnabled(PlaybackButton.Stop, false);
                    _playbackInfo.Enabled = true;
                    break;
            }
        }

        private void StopAudioAndResetControls()
        {
            try
            {
                _audio.Stop();
                _playControl.SetButtonEnabled(PlaybackButton.All, true);
                UpdatePlaybackLocation();
            }
            catch (Exception ex)
            {
                _playbackInfo.PlaybackError = Resources.VOICEMAIL_PCSPEAKER_ERROR;
                _playbackInfo.ShowPlaybackError = true;
                TraceTopics.Voicemail.exception(ex);
            }
        }

        private void StopPlayToNumberCompleted(object sender, AsyncCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                TraceTopics.Voicemail.warning("Error received from VoicemailMessage.StopPlayToNumberCompleted, continuing");
            }
        }

        private void StopPlayToStationCompleted(object sender, AsyncCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                TraceTopics.Voicemail.warning("Error received from VoicemailMessage.StopPlayToStationCompleted, continuing");
            }
        }

        private void OnPlayControlRewind(object sender, EventArgs e)
        {
            OnPlayControlRewindAction();
        }

        private void OnPlayControlRewindAction()
        {
            if (InvokeRequired)
            {
                MethodInvoker del = OnPlayControlRewindAction;
                BeginInvoke(del);
            }
            else
            {
                if ((_audio != null) && (_playbackInfo.SelectedPlaybackDevice == PlaybackDevice.PCSpeakers))
                {
                    _audio.CurrentPosition = _audio.CurrentPosition - (_audio.BytesPerSecond * REWIND_FASTFORWARD_SECONDS);
                    UpdatePlaybackLocation();
                }
            }
        }

        private void OnPlayControlFastForward(object sender, EventArgs e)
        {
            OnPlayControlFastForwardAction();
        }

        private void OnPlayControlFastForwardAction()
        {
            if (InvokeRequired)
            {
                MethodInvoker del = OnPlayControlFastForwardAction;
                BeginInvoke(del);
            }
            else
            {
                if ((_audio != null) && (_playbackInfo.SelectedPlaybackDevice == PlaybackDevice.PCSpeakers))
                {
                    _audio.CurrentPosition = _audio.CurrentPosition + (_audio.BytesPerSecond * REWIND_FASTFORWARD_SECONDS);
                    UpdatePlaybackLocation();
                }
            }
        }

        private void OnPlayControlVolumeChanged(object sender, EventArgs e)
        {
            OnPlayControlVolumeChangedAction();
        }

        private void OnPlayControlVolumeChangedAction()
        {
            if (InvokeRequired)
            {
                MethodInvoker del = OnPlayControlVolumeChangedAction;
                BeginInvoke(del);
            }
            else
            {
                if ((_audio != null) && (_playbackInfo.SelectedPlaybackDevice == PlaybackDevice.PCSpeakers))
                {
                    _audio.Volume = _playControl.Volume;
                }
            }
        }

        private void OnPlayControlSpeedChanged(object sender, EventArgs e)
        {
            OnPlayControlSpeedChangedAction();
        }

        private void OnPlayControlSpeedChangedAction()
        {
            if (InvokeRequired)
            {
                MethodInvoker del = OnPlayControlSpeedChangedAction;
                BeginInvoke(del);
            }
            else
            {
                if ((_audio != null) && (_playbackInfo.SelectedPlaybackDevice == PlaybackDevice.PCSpeakers))
                {
                    switch (_playControl.Speed)
                    {
                        case SpeedLevel.Normal:
                            _audio.Frequency = 0;
                            break;
                        case SpeedLevel.Half:
                            _audio.Frequency = -50;
                            break;
                        case SpeedLevel.Faster:
                            _audio.Frequency = 50;
                            break;
                        case SpeedLevel.Fastest:
                            _audio.Frequency = 100;
                            break;
                    }
                }
            }
        }

        private void OnLocationPanelVisibleChanged(object sender, EventArgs e)
        {
            // When locationPanel is hidden/shown, adjust the entire size of this control to compensate
            if (_locationPanel.Visible)
            {
                if (Height == _initialHeight - _locationPanel.Height)
                {
                    Height = Height + _locationPanel.Height;
                }
            }
            else
            {
                if (Height == _initialHeight)
                {
                    Height = Height - _locationPanel.Height;
                }
            }
        }

        private void VMPlayer_Load(object sender, EventArgs e)
        {
            _initialHeight = Height;
        }

        private void OnRemovingLocalFile(object sender, VoicemailAttachmentContainerEventArgs e)
        {
            if (e.VoicemailAttachmentContainer.Equals(_container))
            {
                // File is no longer available, refresh controls
                ResetUIForLocalPlayback();
            }
        }

        private void ResetUIForLocalPlayback()
        {
            ClearAudio();
            _playbackInfo.SpecialMessage = Resources.VOICEMAIL_DOWNLOAD_REQUIRED;
            _playbackInfo.ShowSpecialMessage = (_container != null) && (_audio == null);
            _playbackInfo.Enabled = true;
            AdjustPlayControls();
        }

        private delegate void OnDownloadCompleteDel(object sender, DownloadEventArgs e);
        private void OnDownloadCompleted(object sender, DownloadEventArgs e)
        {
            if (InvokeRequired)
            {
                OnDownloadCompleteDel del = OnDownloadCompleted;
                BeginInvoke(del, new[] { sender, e });
            }
            
            if (e.VoicemailAttachmentContainer.Equals(_container) && (_downloadManager[_container].Length > 0))
            {
                // Update the UI
                ClearAudio();
                SetupAudio();
                _playbackInfo.SpecialMessage = Resources.VOICEMAIL_DOWNLOAD_REQUIRED;
                _playbackInfo.ShowSpecialMessage = (_container != null) && (_audio == null);
                AdjustPlayControls();

                if (_playAutomatically && (_audio != null))
                {
                    // Begin playing the audio
                    try
                    {
                        _audio.Play();
                    }
                    catch (Exception ex)
                    {
                        _playbackInfo.PlaybackError = Resources.VOICEMAIL_PCSPEAKER_ERROR;
                        _playbackInfo.ShowPlaybackError = true;
                        TraceTopics.Voicemail.exception(ex);
                    }
                    
                    _playAutomatically = false;
                }
            }
            else if (e.Error != null)
            {
                // Download failed, show message to user
                ResetUIForLocalPlayback();
                LocalizedMessageBox.Show(String.Format("{0}{1}{2}", Resources.VOICEMAIL_PLAYBACK_ERROR, Environment.NewLine, Resources.VOICEMAIL_PLAYBACK_ERROR_SOLUTION), Resources.VOICEMAIL_PLAYER_ERROR_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private delegate void OnAudioPausingDel();
        private void OnAudioPausing(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                OnAudioPausingDel del = OnAudioPausingAction;
                BeginInvoke(del);
            }
            else
            {
                OnAudioPausingAction();
            }
        }

        private void OnAudioPausingAction()
        {
            StopTimer();
            _timerStarted = false;

            _playControl.Playing = _audio.Playing;
        }

        private delegate void OnAudioStartingDel();
        private void OnAudioStarting(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                OnAudioStartingDel del = OnAudioStartingAction;
                BeginInvoke(del);
            }
            else
            {
                OnAudioStartingAction();
            }
        }

        private void OnAudioStartingAction()
        {
            StartTimer();
            _timerStarted = true;

            _playControl.Playing = _audio.Playing;
            _playbackInfo.Enabled = !_audio.Playing;
        }

        private delegate void OnAudioStoppingDel();
        private void OnAudioStopping(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                OnAudioStoppingDel del = OnAudioStoppingAction;
                BeginInvoke(del);
            }
            else
            {
                OnAudioStoppingAction();
            }
        }

        private void OnAudioStoppingAction()
        {
            StopTimer();
            _timerStarted = false;

            _playControl.Playing = (_audio != null) ? _audio.Playing : false;
            _playbackInfo.Enabled = (_audio != null) ? !_audio.Playing : true;
        }

        private void StartTimer()
        {
            _updateTimer.Start();
            UpdatePlaybackLocation();
        }

        private void StopTimer()
        {
            _updateTimer.Stop();
            UpdatePlaybackLocation();
        }

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            UpdatePlaybackLocation();
        }

        private void UpdatePlaybackLocation()
        {
            if (_audio != null)
            {
                try
                {
                    // Update location
                    _playbackLocation.ValueChanged -= OnPlaybackLocationChanged;
                    double percent = (_audio.CurrentPosition / (double)_audio.Duration) * 100;
                    int integerPercent = Convert.ToInt32(percent);
                    int finalPercent = Math.Max(Math.Min(integerPercent, 100), 0);
                    _playbackLocation.Value = finalPercent;
                    _playbackLocation.ValueChanged += OnPlaybackLocationChanged;
                }
                catch (Exception ex)
                {
                    TraceTopics.Voicemail.exception(ex);
#if DEBUG
                    throw;
#endif
                }
            }
        }

        private void OnPlaybackLocationChanged(object sender, EventArgs e)
        {
            try
            {
                if (_audio != null)
                {
                    // Translate percent to position
                    double percent = (double)_playbackLocation.Value / 100;
                    _audio.CurrentPosition = Convert.ToInt64(_audio.Duration * percent);
                }
            }
            catch (Exception ex)
            {
                TraceTopics.Voicemail.exception(ex);
#if DEBUG
                throw;
#endif
            }
        }

        private void OnPlaybackLocationMouseDown(object sender, MouseEventArgs e)
        {
            // Stop timer, if it is enabled
            if (_updateTimer.Enabled)
            {
                StopTimer();
            }
        }

        private void OnPlaybackLocationMouseUp(object sender, MouseEventArgs e)
        {
            if (_timerStarted)
            {
                StartTimer();
            }
        }

        private void OnMessagePlaybackComplete(object sender, EventArgs e)
        {
            _userInterfaceContext.ExecuteAsync(() =>
                                              {
                                                  _playControl.SetButtonEnabled(PlaybackButton.Play, true);
                                                  _playControl.SetButtonEnabled(PlaybackButton.Stop, false);
                                                  _playbackInfo.Enabled = true;
                                              });
        }
    }
}
