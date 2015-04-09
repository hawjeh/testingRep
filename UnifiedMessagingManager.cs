using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using ININ.IceLib.Connection;
using ININ.IceLib.Internal;
using ININ.ThinBridge;

namespace ININ.IceLib.UnifiedMessaging
{
    /// <summary>
    /// Provides access to the UnifiedMessaging namespace.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Use the <see cref="UnifiedMessagingManager"/> class to access the functionality
    /// found in the <see cref="ININ.IceLib.UnifiedMessaging"/> namespace.
    /// </para>
    /// <para>
    /// The <see cref="ININ.IceLib.UnifiedMessaging"/> namespace includes all functionality for manipulating
    /// Fax and voicemail data managed by an IC server.  Most features provided in
    /// the <see cref="UnifiedMessagingManager"/> are available in synchronous and
    /// asynchronous versions, allowing you to choose the model that best suits your
    /// needs.
    /// </para>
    /// <br/>
    /// <example>
    /// All "manager" classes found in the IceLib library are designed as
    /// singletons. To begin working with any of the functionality provided in
    /// the <see cref="UnifiedMessagingManager"/> you must obtain the instance
    /// through a call to <see cref="GetInstance"/>.<br/><br/>
    /// <code lang="C#">
    /// Session session = new Session();
    /// session.Connect(...);
    /// UnifiedMessagingManager unifiedMessagingManager = UnifiedMessagingManager.GetInstance(session);
    /// </code>
    /// </example>
    /// </remarks>
    /// <doccompleted/>
    public class UnifiedMessagingManager
    {
        private static readonly ManagedWeakReferenceCache<Session, UnifiedMessagingManager> _UnifiedMessagingManagerCache =
            new ManagedWeakReferenceCache<Session, UnifiedMessagingManager>(TraceTopics.UnifiedMessaging);
        private readonly Dictionary<string, FaxTracker> _FaxTrackers = new Dictionary<string, FaxTracker>(); // remoteFaxFile => callback & local file
        private readonly VoiceMailCache _VMCache;
        private readonly FaxCache _FaxCache;
        private bool _VoicemailWaiting;
        private int _WatchingVMWaitingRefCount;

        private readonly Session _Session;

        private const string _UidRegexUidGroup = "UID";
        private static readonly string _UidRegex = String.Format(CultureInfo.InvariantCulture, "uid=(?<{0}>[0-9]+)", _UidRegexUidGroup);
        private static readonly XmlSerializer _XmlFaxListSerializer;
        private static readonly XmlSerializer _XmlVoicemailListSerializer;

        private readonly Dictionary<string, VoicemailPlayEventTarget> _VoicemailPlayEventTargets = new Dictionary<string, VoicemailPlayEventTarget>();

        #region Async Support

        private readonly AsyncTaskTracker<AsyncUnifiedMessagingManagerState> _AsyncTaskTracker;

        /// <summary>
        /// Private class to store the internal state of a Unified Messaging operation during asynchronous processing.
        /// </summary>
        private class AsyncUnifiedMessagingManagerState : AsyncTaskState
        {
            // Callback values.
            private readonly AsyncCompletedEventHandler _CompletedCallback;

            public AsyncCompletedEventHandler CompletedCallback
            {
                get { return (_CompletedCallback); }
            }

            private readonly EventHandler<AsyncSendFaxCompletedEventArgs> _SendFaxCompletedCallback;

            public EventHandler<AsyncSendFaxCompletedEventArgs> SendFaxCompletedCallback
            {
                get { return (_SendFaxCompletedCallback); }
            }

            private readonly EventHandler<AsyncGetFaxServerSettingsCompletedEventArgs> _GetFaxServerSettingsCompletedCallback;

            public EventHandler<AsyncGetFaxServerSettingsCompletedEventArgs> GetFaxServerSettingsCompletedCallback
            {
                get { return (_GetFaxServerSettingsCompletedCallback); }
            }

            private readonly EventHandler<AsyncGetFaxPropertiesEventArgs> _GetFaxPropertiesCompletedCallback;

            public EventHandler<AsyncGetFaxPropertiesEventArgs> GetFaxPropertiesCompletedCallback
            {
                get { return (_GetFaxPropertiesCompletedCallback); }
            }

            // Input values.
            private readonly bool _Enabled;

            public bool Enabled
            {
                get { return (_Enabled); }
            }

            private FaxResult _FaxResult;

            public FaxResult FaxResult
            {
                get { return (_FaxResult); }
                set { _FaxResult = value; }
            }

            private string _FileName;

            public string FileName
            {
                get { return (_FileName); }
                set { _FileName = value; }
            }

            private int _ActualMessageCount;

            public int ActualMessageCount
            {
                get { return (_ActualMessageCount); }
                set { this._ActualMessageCount = value; }
            }

            private int _MaxMessages;

            public int MaxMessages
            {
                get { return (this._MaxMessages); }
                set { this._MaxMessages = value; }
            }

            private UInt32 _EnvelopeId;

            public UInt32 EnvelopeId
            {
                get { return (this._EnvelopeId); }
                set { this._EnvelopeId = value; }
            }

            private FaxServerSettings _FaxServerSettings;

            public FaxServerSettings FaxServerSettings
            {
                get { return (this._FaxServerSettings); }
                set { this._FaxServerSettings = value; }
            }

            private FaxEnvelopeProperties _FaxEnvelopeProperties;

            public FaxEnvelopeProperties FaxEnvelopeProperties
            {
                get { return (this._FaxEnvelopeProperties); }
                set { this._FaxEnvelopeProperties = value; }
            }

            public AsyncUnifiedMessagingManagerState(int actualMessageCount, long taskId, AsyncCompletedEventHandler completedCallback, object userState)
                    : base(taskId, userState)
            {
                _MaxMessages = actualMessageCount;
                _ActualMessageCount = actualMessageCount;
                _CompletedCallback = completedCallback;
            }

            public AsyncUnifiedMessagingManagerState(UInt32 envelopeId, long taskId, AsyncCompletedEventHandler completedCallback, object userState)
                : base(taskId, userState)
            {
                _EnvelopeId = envelopeId;
                _CompletedCallback = completedCallback;
            }

            public AsyncUnifiedMessagingManagerState(string fileName, long taskId, EventHandler<AsyncSendFaxCompletedEventArgs> completedCallback, object userState)
                    : base(taskId, userState)
            {
                _FileName = fileName;
                _SendFaxCompletedCallback = completedCallback;
            }

            public AsyncUnifiedMessagingManagerState(FaxResult faxResult, long taskId, AsyncCompletedEventHandler completedCallback, object userState)
                    : base(taskId, userState)
            {
                _FaxResult = faxResult;
                _CompletedCallback = completedCallback;
            }

            public AsyncUnifiedMessagingManagerState(bool enabled, long taskId, AsyncCompletedEventHandler completedCallback, object userState)
                    : base(taskId, userState)
            {
                _Enabled = enabled;
                _CompletedCallback = completedCallback;
            }

            public AsyncUnifiedMessagingManagerState(long taskId, AsyncCompletedEventHandler completedCallback, object userState)
                    : base(taskId, userState)
            {
                _CompletedCallback = completedCallback;
            }

            public AsyncUnifiedMessagingManagerState(long taskId, EventHandler<AsyncGetFaxServerSettingsCompletedEventArgs> completedCallback, object userState)
                    : base(taskId, userState)
            {
                _GetFaxServerSettingsCompletedCallback = completedCallback;
            }

            public AsyncUnifiedMessagingManagerState(UInt32 envelopeId, long taskId, EventHandler<AsyncGetFaxPropertiesEventArgs> completedCallback, object userState)
                    : base(taskId, userState)
            {
                _EnvelopeId = envelopeId;
                _GetFaxPropertiesCompletedCallback = completedCallback;
            }
        }

        #endregion

        static UnifiedMessagingManager()
        {
            _XmlFaxListSerializer = new XmlFaxListSerializer();
            _XmlFaxListSerializer.UnknownNode += XML_UnknownNode;
            _XmlFaxListSerializer.UnknownAttribute += XML_UnknownAttribute;

            _XmlVoicemailListSerializer = new XmlVoicemailListSerializer();
            _XmlVoicemailListSerializer.UnknownNode += XML_UnknownNode;
            _XmlVoicemailListSerializer.UnknownAttribute += XML_UnknownAttribute;
        }

        private UnifiedMessagingManager()
        {
            _VMCache = new VoiceMailCache(this);
            _FaxCache = new FaxCache(this);
        }

        private UnifiedMessagingManager(Session session)
                : this()
        {
            _Session = session;
            _Session.RegisterInternalConnectionStateChangedHandler(OnInternalConnectionStateChanged);

            _AsyncTaskTracker = new AsyncTaskTracker<AsyncUnifiedMessagingManagerState>(session);

            RegisterEventHandlers();
        }
        
        private void OnInternalConnectionStateChanged(object sender, ConnectionStateChangedEventArgs e)
        {
            if (e.State != ConnectionState.Down) return;

            // Clear out the mapping if we've lost connection.
            lock (_VoicemailPlayEventTargets)
            {
                _VoicemailPlayEventTargets.Clear();
            }
        }

        /// <summary>
        /// Gets a UnifiedMessagingManager.
        /// </summary>
        /// <param name="session">The Session with which it is associated.</param>
        /// <returns>The UnifiedMessagingManager.</returns>
        /// <exception cref="System.ArgumentNullException">A parameter is <see langword="null"/>.</exception>
        public static UnifiedMessagingManager GetInstance(Session session)
        {
            return _UnifiedMessagingManagerCache.GetOrCreate(session, () => new UnifiedMessagingManager(session));
        }

        /// <summary>
        /// Gets the Session with which this UnifiedMessagingManager is associated.
        /// </summary>
        /// <value>The session.</value>
        public Session Session
        {
            get { return _Session; }
        }

        private void RegisterEventHandlers()
        {
            try
            {
                _Session.RegisterMessageHandler(InternalThinSessionEventIds.eThinSessionEvt_FaxMonitorUpdate, HandleFaxMonitorUpdate);
                _Session.RegisterMessageHandler(InternalThinSessionEventIds.eThinSessionEvt_FaxSubmitResult, HandleSendFaxResult);
                _Session.RegisterMessageHandler(InternalThinSessionEventIds.eThinSessionEvt_MessageLightEvent, HandleMessageLight);
                _Session.RegisterMessageHandler(InternalThinSessionEventIds.eThinSessionEvt_VoiceMessageList, HandleVoiceMessageList);
                _Session.RegisterMessageHandler(InternalThinSessionEventIds.eThinSessionEvt_VoiceMessagePlayResult, HandleMessagePlayResult);
                _Session.RegisterMessageHandler(InternalThinSessionEventIds.eThinSessionEvt_UMMessageList, HandleFaxMessageList);
            }
            catch (Exception ex)
            {
                TraceTopics.UnifiedMessaging.exception(ex);
            }
        }

        private void DeregisterEventHandlers()
        {
            try
            {
                _Session.DeregisterMessageHandler(InternalThinSessionEventIds.eThinSessionEvt_FaxMonitorUpdate, HandleFaxMonitorUpdate);
                _Session.DeregisterMessageHandler(InternalThinSessionEventIds.eThinSessionEvt_FaxSubmitResult, HandleSendFaxResult);
                _Session.DeregisterMessageHandler(InternalThinSessionEventIds.eThinSessionEvt_MessageLightEvent, HandleMessageLight);
                _Session.DeregisterMessageHandler(InternalThinSessionEventIds.eThinSessionEvt_VoiceMessageList, HandleVoiceMessageList);
                _Session.DeregisterMessageHandler(InternalThinSessionEventIds.eThinSessionEvt_VoiceMessagePlayResult, HandleMessagePlayResult);
                _Session.DeregisterMessageHandler(InternalThinSessionEventIds.eThinSessionEvt_UMMessageList, HandleFaxMessageList);
            }
            catch (Exception ex)
            {
                TraceTopics.UnifiedMessaging.exception(ex);
            }
        }

        #region Fax Creation API

        /// <summary>
        /// Occurs when monitoring is turned on and a Fax event is detected.
        /// </summary>
        /// <remarks>
        /// This event will only occur if faxes are being monitored. To start monitoring
        /// faxes call either the <see cref="EnableFaxMonitoring"/> method or
        /// the <see cref="EnableFaxMonitoringAsync"/> method.
        /// <note>All event handlers should be added before calling
        /// <see cref="EnableFaxMonitoring"/> or 
        /// <see cref="EnableFaxMonitoringAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// </remarks>
        public event EventHandler<FaxMonitorUpdateEventArgs> FaxMonitorUpdate;

        #region EnableFaxMonitoring

        private void EnableFaxMonitoringImpl(bool enabled)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                try
                {
                    Message request = _Session.CreateMessage(InternalThinManagerEventIds.eThinManagerReq_FaxMonitor);
                    request.Payload.Writer.Write(enabled);
                    _Session.SendMessageAndThrowIfError(request);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues an synchronous request to enable or disable monitoring of Fax server events.
        /// </summary>
        /// <param name="enabled">Should monitoring be turned on.</param>
        /// <ConnectionExceptions />
        public void EnableFaxMonitoring(bool enabled)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked. enabled={}", enabled);

                try
                {
                    EnableFaxMonitoringImpl(enabled);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues an asynchronous request to enable or disable monitoring of Fax server events.
        /// </summary>
        /// <param name="enabled">Should monitoring be turned on.</param>
        /// <param name="completedCallback">The callback to invoke when the asynchronous operation completes.</param>
        /// <param name="userState"><ininAsyncStateParam /></param>
        /// <remarks><ininAsyncMethodNote /></remarks>
        public void EnableFaxMonitoringAsync(bool enabled, AsyncCompletedEventHandler completedCallback, object userState)
        {
            long taskId = _AsyncTaskTracker.CreateTaskId();
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(taskId))
            {
                TraceTopics.UnifiedMessaging.status("Invoked. enabled={}", enabled);

                try
                {
                    AsyncUnifiedMessagingManagerState asyncState = new AsyncUnifiedMessagingManagerState(enabled, taskId, completedCallback, userState);
                    _AsyncTaskTracker.PerformTask(asyncState, EnableFaxMonitoringAsyncPerformTask, EnableFaxMonitoringAsyncCompleted);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async operation (called on an arbitrary worker thread).
        private void EnableFaxMonitoringAsyncPerformTask(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked. enabled={}", asyncState.Enabled);

                try
                {
                    EnableFaxMonitoringImpl(asyncState.Enabled);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async completion (called on the UI thread).
        private void EnableFaxMonitoringAsyncCompleted(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            {
                AsyncCompletedEventArgs e = new AsyncCompletedEventArgs(
                        asyncState.Exception,
                        asyncState.Cancelled,
                        asyncState.UserState);
                EventHelper.SafeRaise(asyncState.CompletedCallback, this, e, TraceTopics.UnifiedMessaging, asyncState.Enabled);
            }
        }

        #endregion // EnableFaxMonitoring

        private class SendFaxWaitState : IntermediateWaitState
        {
            private readonly String _FaxFile;
            private readonly String _RemoteFile;
            private readonly OperationDelegate _Operation;

            public SendFaxWaitState(string faxFile, string remoteFile, OperationDelegate operation, AsyncCallback callback, object state)
                    : base(callback, state)
            {
                _FaxFile = faxFile;
                _RemoteFile = remoteFile;
                _Operation = operation;
            }

            public String FaxFile
            {
                get { return _FaxFile; }
            }

            public String RemoteFile
            {
                get { return _RemoteFile; }
            }

            public OperationDelegate Operation
            {
                get { return _Operation; }
            }
        }

        private class FaxTracker
        {
            private readonly String _FileName;
            private UInt32 _FaxId;
            private SendFaxWaitState _WaitState;

            public FaxTracker(String fileName)
            {
                _FileName = fileName;
            }

            public String FileName
            {
                get { return _FileName; }
            }

            public UInt32 FaxId
            {
                get { return _FaxId; }
                set { _FaxId = value; }
            }

            public SendFaxWaitState WaitState
            {
                get { return _WaitState; }
                set { _WaitState = value; }
            }
        }

        #region SendFax

        private FaxResult SendFaxImpl(string fileName)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                _Session.CheckSession();

                Guard.ArgumentNotNull(fileName, "fileName");
                if (fileName.Length == 0)
                    throw new ArgumentException(Localization.LoadString(_Session, "EMPTY_STRING_ERROR"), "fileName");
                if (File.Exists(fileName) == false)
                    throw new ArgumentException(Localization.LoadString(_Session, "FILE_NOT_FOUND_ERROR"), "fileName");

                try
                {
                    return (SendFaxInternal(fileName));
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>Issues a synchronous request to send a Fax.</summary>
        /// <param name="fileName">The file name of the Fax to send.</param>
        /// <returns>A <see cref="FaxResult"/> object with the results of the SendFax method.</returns>
        /// <exception cref="System.ArgumentException">If a <c>fileName</c> of length zero.</exception>
        /// <exception cref="System.ArgumentNullException">A parameter is <see langword="null"/>.</exception>
        /// <ConnectionExceptions />
        /// <remarks>
        /// <para>
        /// Use the <see cref="FaxFile"/> class to create a .TIF or .i3f
        /// formatted Fax file in the file system before sending.
        /// </para>
        /// <para><note>
        /// Most fax terminals can't handle widths larger than 1728 pixels. If the width
        /// of the <see cref="System.Drawing.Image"/> will exceed 1728 pixels, it is recommended to
        /// store the image in portrait orientation and set the <see cref="FaxPageAttributes"/>
        /// instance to indicate that the image has been rotated.<br/>
        /// This can also be handled by specifying <see cref="FaxImageSettings.FitToPage"/> when adding or updating pages.
        /// This will cause the image to be scaled to fit on the page.  If <see cref="FaxImageSettings.FitToPage"/> is not
        /// specified, no resizing will be done and the image will be cropped.
        /// </note></para>
        /// </remarks>
        /// <seealso cref="FaxFile"/>
        /// <ConnectionExceptions />
        public FaxResult SendFax(string fileName)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    Guard.ArgumentNotNull(fileName, "fileName");

                    if (fileName.Length == 0)
                        throw new ArgumentException(Localization.LoadString(_Session, "EMPTY_STRING_ERROR"), "fileName");

                    return SendFaxImpl(fileName);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues an asynchronous request to send a Fax.
        /// </summary>
        /// <param name="fileName">The file name of the Fax to send.</param>
        /// <param name="completedCallback">The callback to invoke when the asynchronous operation completes.</param>
        /// <param name="userState"><ininAsyncStateParam /></param>
        /// <exception cref="System.ArgumentNullException">A parameter is <see langword="null"/>.</exception>
        /// <exception cref="System.ArgumentException">If a <c>fileName</c> of length zero.</exception>
        /// <remarks>
        /// <para>
        /// Use the <see cref="FaxFile"/> class to create a .TIF or .i3f
        /// formatted Fax file in the file system before sending.
        /// </para>
        /// <para><note>
        /// Most fax terminals can't handle widths larger than 1728 pixels. If the width
        /// of the <see cref="System.Drawing.Image"/> will exceed 1728 pixels, it is recommended to
        /// store the image in portrait orientation and set the <see cref="FaxPageAttributes"/>
        /// instance to indicate that the image has been rotated.<br/>
        /// This can also be handled by specifying <see cref="FaxImageSettings.FitToPage"/> when adding or updating pages.
        /// This will cause the image to be scaled to fit on the page.  If <see cref="FaxImageSettings.FitToPage"/> is not
        /// specified, no resizing will be done and the image will be cropped.
        /// </note></para>
        /// <ininAsyncMethodNote />
        /// </remarks>
        /// <seealso cref="FaxFile"/>
        public void SendFaxAsync(string fileName, EventHandler<AsyncSendFaxCompletedEventArgs> completedCallback, object userState)
        {
            long taskId = _AsyncTaskTracker.CreateTaskId();
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(taskId))
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    Guard.ArgumentNotNull(fileName, "fileName");

                    if (fileName.Length == 0)
                        throw new ArgumentException(Localization.LoadString(_Session, "EMPTY_STRING_ERROR"), "fileName");

                    AsyncUnifiedMessagingManagerState asyncState = new AsyncUnifiedMessagingManagerState(fileName, taskId, completedCallback, userState);
                    _AsyncTaskTracker.PerformTask(asyncState, SendFaxAsyncPerformTask, SendFaxAsyncCompleted);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async operation (called on an arbitrary worker thread).
        private void SendFaxAsyncPerformTask(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    asyncState.FaxResult = SendFaxImpl(asyncState.FileName);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async completion (called on the UI thread).
        private void SendFaxAsyncCompleted(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            {
                AsyncSendFaxCompletedEventArgs e = new AsyncSendFaxCompletedEventArgs(
                        asyncState.FaxResult,
                        asyncState.Exception,
                        asyncState.Cancelled,
                        asyncState.UserState);
                EventHelper.SafeRaise(asyncState.SendFaxCompletedCallback, this, e, TraceTopics.UnifiedMessaging);
            }
        }

        #endregion // SendFax

        #region Fax Send Implementation

        /// <summary>
        /// Sends a Fax.
        /// </summary>
        /// <param name="fileName">The full local pathname of the file to submit for Faxing.</param>
        /// <returns>The <see cref="FaxResult"/> that was returned.</returns>
        private FaxResult SendFaxInternal(string fileName)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                try
                {
                    TraceTopics.UnifiedMessaging.status("Invoked. fileName={}", fileName);

                    IAsyncResult asyncResult = BeginSendFax(fileName, null, null);
                    return EndSendFax(asyncResult);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Begins an asynchronous request to send a Fax.
        /// </summary>
        /// <param name="fileName">The full local pathname of the file to submit for Faxing.</param>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state"><ininAsyncStateParam /></param>
        /// <returns>
        /// An IAsyncResult that references the asynchronous operation.
        /// </returns>
        /// <exception cref="System.ArgumentNullException">A parameter is <see langword="null"/>.</exception>
        private IAsyncResult BeginSendFax(string fileName, AsyncCallback callback, object state)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                try
                {
                    TraceTopics.UnifiedMessaging.status("Invoked. fileName={}", fileName);

                    Guard.ArgumentNotNull(fileName, "fileName");

                    string remoteFileName = RemoteFileHelper.NewRemoteFileName(ServerFileType.Fax);

                    return BeginSendFaxInternal(fileName, remoteFileName, callback, state);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Ends a pending asynchronous request to send a Fax.
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that stores state information and any user defined 
        /// data for this asynchronous operation.</param>
        /// <returns>The <see cref="FaxResult"/> that was returned.</returns>
        /// <exception cref="System.ArgumentNullException">A parameter is <see langword="null"/>.</exception>
        private FaxResult EndSendFax(IAsyncResult asyncResult)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                try
                {
                    TraceTopics.UnifiedMessaging.status("Invoked.");

                    Guard.ArgumentNotNull(asyncResult, "asyncResult");

                    return EndSendFaxInternal(asyncResult);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        #region GetFaxServerSettings

        private FaxServerSettings GetFaxServerSettingsImpl()
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                try
                {
                    Message request = _Session.CreateMessage(InternalThinManagerEventIds.eThinManagerReq_FaxGetServerSettings);
                    Message response = _Session.SendMessageAndThrowIfError(request);

                    return FaxServerSettings.Read(response.Payload.Reader);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues a synchronous request to retrieve the Fax configuration settings on the IC server.
        /// </summary>
        /// <returns>A <see cref="ININ.IceLib.UnifiedMessaging.FaxServerSettings"/> object with the Fax 
        /// configuration settings on the IC server.</returns>
        /// <ConnectionExceptions />
        public FaxServerSettings GetFaxServerSettings()
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    return GetFaxServerSettingsImpl();
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues an asynchronous request to retrieve the Fax configuration settings on the IC server.
        /// </summary>
        /// <param name="completedCallback">The callback to invoke when the asynchronous operation completes.</param>
        /// <param name="userState"><ininAsyncStateParam /></param>
        /// <remarks><ininAsyncMethodNote /></remarks>
        public void GetFaxServerSettingsAsync(EventHandler<AsyncGetFaxServerSettingsCompletedEventArgs> completedCallback, object userState)
        {
            long taskId = _AsyncTaskTracker.CreateTaskId();
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(taskId))
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    AsyncUnifiedMessagingManagerState asyncState = new AsyncUnifiedMessagingManagerState(taskId, completedCallback, userState);
                    _AsyncTaskTracker.PerformTask(asyncState, GetFaxServerSettingsAsyncPerformTask, GetFaxServerSettingsAsyncCompleted);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async operation (called on an arbitrary worker thread).
        private void GetFaxServerSettingsAsyncPerformTask(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    asyncState.FaxServerSettings = GetFaxServerSettingsImpl();
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async completion (called on the UI thread).
        private void GetFaxServerSettingsAsyncCompleted(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            {
                AsyncGetFaxServerSettingsCompletedEventArgs e = new AsyncGetFaxServerSettingsCompletedEventArgs(
                        asyncState.FaxServerSettings,
                        asyncState.Exception,
                        asyncState.Cancelled,
                        asyncState.UserState);
                EventHelper.SafeRaise(asyncState.GetFaxServerSettingsCompletedCallback, this, e, TraceTopics.UnifiedMessaging);
            }
        }

        #endregion // GetFaxServerSettings

        private delegate void OperationDelegate(string fileName, string remoteFileName);

        private IAsyncResult BeginSendFaxInternal(string fileName, string remoteFileName, AsyncCallback callback, object state)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                _Session.CheckSession();
                Guard.ArgumentNotNull(fileName, "fileName");
                if (fileName.Length == 0)
                    throw new ArgumentException(Localization.LoadString(_Session, "EMPTY_STRING_ERROR"), "fileName");
                if (File.Exists(fileName) == false)
                    throw new ArgumentException(Localization.LoadString(_Session, "FILE_NOT_FOUND_ERROR"), "fileName");

                try
                {
                    OperationDelegate putAndSendDelegate = new OperationDelegate(PutAndSendOperation);

                    SendFaxWaitState waitState = new SendFaxWaitState(fileName, remoteFileName, putAndSendDelegate, callback, state);

                    // Track the fax filename along with the appropriate callback,
                    // to hook up the event from Session Manager with the callback.
                    FaxTracker tracker = GetFaxTracker(waitState.RemoteFile, true);
                    tracker.WaitState = waitState;

                    waitState.IntermediateResult = putAndSendDelegate.BeginInvoke(fileName, remoteFileName, new AsyncCallback(PutAndSendIntermediateCallback), waitState);

                    return waitState;
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        private void PutAndSendOperation(string fileName, string remoteFileName)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                // (1) Perform the PutFile operation to push the Fax file up to the server
                long MaxPacketSize = 2048;
                int packetType = InternalFileTransferPacketType.eFileTransferPacketBegin;
                byte[] buffer = new byte[MaxPacketSize];

                // Pull the information out of the file...
                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                BinaryReader reader = new BinaryReader(fs);

                long remaining = reader.BaseStream.Length;
                reader.BaseStream.Position = 0;

                do
                {
                    int length = reader.Read(buffer, 0, buffer.Length);
                    remaining -= length;
                    if (remaining == 0)
                    {
                        packetType |= InternalFileTransferPacketType.eFileTransferPacketEnd;
                        reader.Close();
                    }

                    Message putRequest = _Session.CreateMessage(InternalThinManagerEventIds.eThinManagerReq_PutUserFile);
                    putRequest.Payload.Writer
                            .Write(remoteFileName)
                            .Write(packetType)
                            .Write(length)
                            .PutData(buffer, length);

                    _Session.SendMessageAndThrowIfError(putRequest);

                    packetType &= ~ InternalFileTransferPacketType.eFileTransferPacketBegin;
                } while (remaining > 0);

                // (2) Now submit the fax to the server for a send
                Message submitRequest = _Session.CreateMessage(InternalThinManagerEventIds.eThinManagerReq_FaxSubmit);
                submitRequest.Payload.Writer.Write(Path.GetFileName(remoteFileName));
                _Session.SendMessageAndThrowIfError(submitRequest);
            }
        }

        private void PutAndSendIntermediateCallback(IAsyncResult asyncResult)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                Guard.ArgumentNotNull(asyncResult, "asyncResult");
                SendFaxWaitState waitState = asyncResult.AsyncState as SendFaxWaitState;
                if (waitState == null)
                    throw new ArgumentException("Mismatched asyncResult object", "asyncResult");

                try
                {
                    waitState.Operation.EndInvoke(asyncResult);
                }
                catch (Exception ex)
                {
                    TraceTopics.UnifiedMessaging.exception(ex);

                    waitState.ExceptionToThrow = ex;

                    // because we hit an exception, go ahead and signal the end of the operation
                    waitState.SignalCompletion();
                }
            }
        }

        private FaxResult EndSendFaxInternal(IAsyncResult asyncResult)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                Guard.ArgumentNotNull(asyncResult, "asyncResult");

                SendFaxWaitState waitState = asyncResult as SendFaxWaitState;
                if (waitState == null)
                    throw new ArgumentException("Mismatched result object", "asyncResult");

                // Block until the result is available
                waitState.WaitForCompletion();

                FaxTracker tracker = GetFaxTracker(waitState.RemoteFile, false);
                if (tracker == null)
                {
                    TraceTopics.UnifiedMessaging.error("Fax not found in tracker: file={}, remote={}", waitState.FaxFile, waitState.RemoteFile);
                    throw new ArgumentException(Localization.LoadString(_Session, "MISMATCHED_RESULT_ERROR"), "asyncResult");
                }

                FaxResult faxResult = new FaxResult(waitState.FaxFile, tracker.FaxId);

                // remove the tracker and signal the end of the fax send operation
                _FaxTrackers.Remove(waitState.RemoteFile);

                TraceTopics.UnifiedMessaging.status("Returning: {}", faxResult);

                return faxResult;
            }
        }

        private void HandleSendFaxResult(object sender, MessageReceivedEventArgs e)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                // Read the name of the 'remote' file that was just sent
                String remoteFileName = e.Message.Payload.Reader.ReadString();

                FaxTracker tracker = GetFaxTracker(remoteFileName, false);
                if (tracker != null)
                {
                    try
                    {
                        // Read the send response
                        int errorCode;
                        string errorDescription;
                      
                        StreamErrorInformation.ReadLastError(e.Message.Payload.Reader, out errorCode, out errorDescription, true);

                        if (errorCode < 0)
                        {
                            tracker.WaitState.ExceptionToThrow = new IceLibException(errorDescription);
                        }

                        tracker.FaxId = e.Message.Payload.Reader.ReadUInt32();

                        // Signal the end of the fax send operation
                        tracker.WaitState.SignalCompletion();
                    }
                    catch (Exception ex)
                    {
                        TraceTopics.UnifiedMessaging.exception(ex);
                        tracker.WaitState.ExceptionToThrow = ex;
                    }
                }
                else
                {
                    TraceTopics.UnifiedMessaging.warning("Fax not found in tracker: remote={}", remoteFileName);
                }
            }
        }

        private FaxTracker GetFaxTracker(string fileName, bool createNew)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                Guard.ArgumentNotNullOrEmptyString(fileName, "fileName");

                FaxTracker tracker = null;
                lock (_FaxTrackers)
                {
                    string pathedFileName = Path.GetFileName(fileName);
                    if (_FaxTrackers.ContainsKey(pathedFileName))
                    {
                        tracker = _FaxTrackers[pathedFileName];
                    }

                    if (tracker == null)
                    {
                        if (createNew)
                        {
                            tracker = new FaxTracker(fileName);
                            _FaxTrackers[Path.GetFileName(fileName)] = tracker;
                        }
                    }
                }

                return tracker;
            }
        }

        #endregion

        #region Cancel Fax

        private void CancelFaxImpl(UInt32 envelopeId)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                try
                {
                    Message request = _Session.CreateMessage(InternalThinManagerEventIds.eThinManagerReq_FaxCancel);
                    request.Payload.Writer.Write(envelopeId);
                    _Session.SendMessageAndThrowIfError(request);
                }
                catch (Exception ex)
                {
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues a synchronous request to cancel a fax.
        /// </summary>
        /// <param name="envelopeId">The envelope ID of the fax.</param>
        /// <exception cref="ArgumentOutOfRangeException">The argument isn't in the expected range.</exception>
        /// <icversion>3.0 SU 1</icversion>
        public void CancelFax(Int64 envelopeId)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    if ((envelopeId < 0) || (envelopeId > UInt32.MaxValue)) throw new ArgumentOutOfRangeException("envelopeId");

                    UInt32 envelopeIdInternal = (UInt32)envelopeId;
                    CancelFaxImpl(envelopeIdInternal);
                }
                catch (Exception ex)
                {
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues an asynchronous request to cancel a fax.
        /// </summary>
        /// <param name="envelopeId">The envelope ID of the fax.</param>
        /// <param name="completedCallback">The callback to invoke when the asynchronous operation completes.</param>
        /// <param name="userState"><ininAsyncStateParam /></param>
        /// <remarks>
        /// <ininAsyncMethodNote />
        /// </remarks>
        /// <exception cref="ArgumentOutOfRangeException">The argument isn't in the expected range.</exception>
        /// <icversion>3.0 SU 1</icversion>
        public void CancelFaxAsync(Int64 envelopeId, AsyncCompletedEventHandler completedCallback, object userState)
        {
            long taskId = _AsyncTaskTracker.CreateTaskId();
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(taskId))
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    if ((envelopeId < 0) || (envelopeId > UInt32.MaxValue)) throw new ArgumentOutOfRangeException("envelopeId");

                    UInt32 envelopeIdInternal = (UInt32)envelopeId;

                    AsyncUnifiedMessagingManagerState asyncState = new AsyncUnifiedMessagingManagerState(envelopeIdInternal, taskId, completedCallback, userState);
                    _AsyncTaskTracker.PerformTask(asyncState, CancelFaxAsyncPerformTask, CancelFaxAsyncCompleted);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async operation (called on an arbitrary worker thread).
        private void CancelFaxAsyncPerformTask(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    CancelFaxImpl(asyncState.EnvelopeId);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async completion (called on the UI thread).
        private void CancelFaxAsyncCompleted(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            {
                AsyncCompletedEventArgs e = new AsyncCompletedEventArgs(asyncState.Exception, asyncState.Cancelled, asyncState.UserState);
                EventHelper.SafeRaise(asyncState.CompletedCallback, this, e, TraceTopics.UnifiedMessaging);
            }
        }

        #endregion

        #region Get Fax Properties

        private FaxEnvelopeProperties GetFaxPropertiesImpl(UInt32 envelopeId)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                FaxEnvelopeProperties faxEnvelopeProperties = null;

                try
                {
                    Message request = _Session.CreateMessage(InternalThinManagerEventIds.eThinManagerReq_FaxViewProperties);
                    request.Payload.Writer.Write(envelopeId);
                    Message response = _Session.SendMessageAndThrowIfError(request);

                    faxEnvelopeProperties = FaxEnvelopeProperties.ReadFromMessage(response.Payload.Reader);

                    return (faxEnvelopeProperties);
                }
                catch (Exception ex)
                {
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues a synchronous request to get the fax envelope properties for a fax in progress.
        /// </summary>
        /// <param name="envelopeId">The envelope ID of the fax.</param>
        /// <returns>The <see cref="ININ.IceLib.UnifiedMessaging.FaxEnvelopeProperties"/> for this fax.</returns>
        /// <exception cref="ArgumentOutOfRangeException">The argument isn't in the expected range.</exception>
        /// <icversion>3.0 SU 1</icversion>
        public FaxEnvelopeProperties GetFaxProperties(Int64 envelopeId)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    if ((envelopeId < 0) || (envelopeId > UInt32.MaxValue)) throw new ArgumentOutOfRangeException("envelopeId");

                    UInt32 envelopeIdInternal = (UInt32)envelopeId;

                    return (GetFaxPropertiesImpl(envelopeIdInternal));
                }
                catch (Exception ex)
                {
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues an asynchronous request to get the envelope properties of a fax in progress.
        /// </summary>
        /// <param name="envelopeId">The envelope ID of the fax for which properties are requested.</param>
        /// <param name="completedCallback">The callback to invoke when the asynchronous operation completes.</param>
        /// <param name="userState"><ininAsyncStateParam /></param>
        /// <remarks>
        /// <ininAsyncMethodNote />
        /// </remarks>
        /// <exception cref="ArgumentOutOfRangeException">The argument isn't in the expected range.</exception>
        /// <icversion>3.0 SU 1</icversion>
        public void GetFaxPropertiesAsync(Int64 envelopeId, EventHandler<AsyncGetFaxPropertiesEventArgs> completedCallback, object userState)
        {
            long taskId = _AsyncTaskTracker.CreateTaskId();
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(taskId))
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    if ((envelopeId < 0) || (envelopeId > UInt32.MaxValue)) throw new ArgumentOutOfRangeException("envelopeId");

                    UInt32 envelopeIdInternal = (UInt32)envelopeId;

                    AsyncUnifiedMessagingManagerState asyncState = new AsyncUnifiedMessagingManagerState(envelopeIdInternal, taskId, completedCallback, userState);
                    _AsyncTaskTracker.PerformTask(asyncState, GetFaxPropertiesAsyncPerformTask, GetFaxPropertiesAsyncCompleted);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async operation (called on an arbitrary worker thread).
        private void GetFaxPropertiesAsyncPerformTask(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    asyncState.FaxEnvelopeProperties = GetFaxPropertiesImpl(asyncState.EnvelopeId);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async completion (called on the UI thread).
        private void GetFaxPropertiesAsyncCompleted(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            {
                AsyncGetFaxPropertiesEventArgs e = new AsyncGetFaxPropertiesEventArgs(asyncState.FaxEnvelopeProperties, asyncState.Exception, asyncState.Cancelled, asyncState.UserState);
                EventHelper.SafeRaise(asyncState.GetFaxPropertiesCompletedCallback, this, e, TraceTopics.UnifiedMessaging);
            }
        }

        #endregion

        private void HandleFaxMonitorUpdate(object sender, MessageReceivedEventArgs e)
        {
            // deserialize the fax status from the notifier reader
            FaxStatus status = FaxStatus.Read(_Session, e.Message.Payload.Reader);
            // prepare the update args to notify the listener of the change
            FaxMonitorUpdateEventArgs eventArgs = new FaxMonitorUpdateEventArgs(status);
            EventHelper.SafeRaise(FaxMonitorUpdate, this, eventArgs, TraceTopics.UnifiedMessaging);
        }

        #endregion

        #region Fax UM API

        #region Fax Events

        /// <summary>
        /// Occurs when a new fax is received.
        /// </summary>
        /// <remarks>
        /// This event will only occur if faxes are being monitored. To start monitoring
        /// faxes call either the <see cref="EnableFaxMonitoring"/> method or
        /// the <see cref="EnableFaxMonitoringAsync"/> method.
        /// <note>All event handlers should be added before calling
        /// <see cref="EnableFaxMonitoring"/> or 
        /// <see cref="EnableFaxMonitoringAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// </remarks>
        public event EventHandler<FaxEventArgs> NewFax;

        /// <summary>
        /// Occurs when a fax is deleted.
        /// </summary>
        /// <remarks>
        /// This event will only occur if faxes are being monitored. To start monitoring
        /// faxes call either the <see cref="EnableFaxMonitoring"/> method or
        /// the <see cref="EnableFaxMonitoringAsync"/> method.
        /// <note>All event handlers should be added before calling
        /// <see cref="EnableFaxMonitoring"/> or 
        /// <see cref="EnableFaxMonitoringAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// </remarks>
        public event EventHandler<FaxEventArgs> FaxDeleted;

        /// <summary>
        /// Occurs when a fax is modified.
        /// </summary>
        /// <remarks>
        /// This event will only occur if faxes are being monitored. To start monitoring
        /// faxes call either the <see cref="EnableFaxMonitoring"/> method or
        /// the <see cref="EnableFaxMonitoringAsync"/> method.
        /// <note>All event handlers should be added before calling
        /// <see cref="EnableFaxMonitoring"/> or 
        /// <see cref="EnableFaxMonitoringAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// </remarks>
        public event EventHandler<FaxEventArgs> FaxUpdated;

        private void OnNewFax(FaxEventArgs e)
        {
            EventHelper.SafeRaise(NewFax, this, e, TraceTopics.UnifiedMessaging);
        }

        private void OnFaxDeleted(FaxEventArgs e)
        {
            EventHelper.SafeRaise(FaxDeleted, this, e, TraceTopics.UnifiedMessaging);
        }

        private void OnFaxUpdated(FaxEventArgs e)
        {
            EventHelper.SafeRaise(FaxUpdated, this, e, TraceTopics.UnifiedMessaging);
        }

        #endregion

        /// <summary>
        /// Gets a <see cref="System.Collections.ObjectModel.ReadOnlyCollection{T}"/> of all FaxMessages.
        /// </summary>
        /// <value>The fax messages.</value>
        public ReadOnlyCollection<FaxMessage> Faxes
        {
            get { return Array.AsReadOnly<FaxMessage>(_FaxCache.GetAll()); }
        }

        #region RefreshFaxCache

        private void RefreshFaxCacheImpl(int maxMessages)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked. maxMessages={}", maxMessages);

                try
                {
                    try
                    {
                        if (maxMessages < -1)
                        {
                            maxMessages = -1;
                        }

                        Message request = _Session.CreateMessage(InternalThinManagerEventIds.eThinManagerReq_GetUMMessageList);
                        request.Payload.Writer.Write("FAX");
                        request.Payload.Writer.Write(maxMessages);
                        _Session.SendMessageAndThrowIfError(request);
                    }
                    catch (Exception ex)
                    {
                        if (((IceLibException)ex).ErrorCode == InternalHRESULTs.E_I3ERR_REQUEST_REJECTED)
                        {
                            throw new RateLimitedException(ex);
                        }

                        throw new IceLibException(String.Format(Localization.LoadString(_Session, "REFRESH_LIST_ERROR"), "Fax"), ex);
                    }
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues a synchronous request to refresh the fax cache.
        /// </summary>
        /// <remarks>
        /// Any subsequent calls to this method or its async version (<see cref="RefreshFaxCacheAsync"/>) within a 5 second period will be rejected with a <see cref="RateLimitedException"/>.
        /// This restriction is per message type, a call to <see cref="RefreshFaxCache"/> followed by a call to <see cref="RefreshVoicemailCache"/> or <see cref="RefreshVoicemailCacheAsync"/> within a 5 second period will be successful.
        /// </remarks>
        /// <param name="maxMessages">The number of messages to refresh, -1 to refresh them all.</param>
        /// <exception cref="IceLibException">There was an error refreshing the list or this method was called more than once within a 5 second period.</exception>
        /// <ConnectionExceptions />
        public void RefreshFaxCache(int maxMessages)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked. maxMessages={}", maxMessages);

                try
                {
                    RefreshFaxCacheImpl(maxMessages);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues an asynchronous request to refresh the fax cache.
        /// </summary>
        /// <remarks>
        /// Any subsequent calls to this method or its synchronous version (<see cref="RefreshFaxCache"/>) within a 5 second period will be rejected with a <see cref="RateLimitedException"/>.
        /// This restriction is per message type, a call to <see cref="RefreshFaxCacheAsync"/> followed by a call to <see cref="RefreshVoicemailCache"/> or <see cref="RefreshVoicemailCacheAsync"/> within a 5 second period will be successful.
        /// </remarks>
        /// <param name="maxMessages">The number of messages to refresh, -1 to refresh them all.</param>
        /// <param name="completedCallback">The callback to invoke when the asynchronous operation completes.</param>
        /// <param name="userState"><ininAsyncStateParam /></param>
        /// <remarks><ininAsyncMethodNote /></remarks>
        public void RefreshFaxCacheAsync(int maxMessages, AsyncCompletedEventHandler completedCallback, object userState)
        {
            long taskId = _AsyncTaskTracker.CreateTaskId();
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(taskId))
            {
                TraceTopics.UnifiedMessaging.status("Invoked. maxMessages={}", maxMessages);

                try
                {
                    AsyncUnifiedMessagingManagerState asyncState = new AsyncUnifiedMessagingManagerState(maxMessages, taskId, completedCallback, userState);
                    _AsyncTaskTracker.PerformTask(asyncState, RefreshFaxCacheAsyncPerformTask, RefreshFaxCacheAsyncCompleted);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async operation (called on an arbitrary worker thread).
        private void RefreshFaxCacheAsyncPerformTask(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked. maxMessages={}", asyncState.MaxMessages);

                try
                {
                    RefreshFaxCacheImpl(asyncState.MaxMessages);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async completion (called on the UI thread).
        private void RefreshFaxCacheAsyncCompleted(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            {
                AsyncCompletedEventArgs e = new AsyncCompletedEventArgs(
                        asyncState.Exception,
                        asyncState.Cancelled,
                        asyncState.UserState);
                EventHelper.SafeRaise(asyncState.CompletedCallback, this, e, TraceTopics.UnifiedMessaging, asyncState.MaxMessages);
            }
        }

        #endregion // RefreshFaxCache

        internal void RemoveFromCache(FaxMessage message)
        {
            _FaxCache.Remove(message);
            OnFaxDeleted(new FaxEventArgs(message));
        }

        #region Fax Message Handlers

        private void HandleFaxMessageList(object sender, MessageReceivedEventArgs e)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                try
                {
                    string sType = e.Message.Payload.Reader.ReadString();
                    // For now, sType will always be "FAX", but in the future it will differentiate between FAX and VOICEMAIL

                    string sXML = e.Message.Payload.Reader.ReadString();

                    StringReader sr = new StringReader(sXML);
                    XmlFaxList list = (XmlFaxList) _XmlFaxListSerializer.Deserialize(sr);

                    List<FaxMessage> messages = new List<FaxMessage>();

                    // Start by adding all items to the collection.  For any
                    // that don't already exist we send a notification, if an existing fax
                    // has changed, send the changed notification
                    foreach (XmlMessage message in list.AllMessages)
                    {
                        FaxMessage msg = new FaxMessage(message, this);
                        if (_FaxCache.Contains(msg))
                        {
                            bool changed, added;
                            _FaxCache.Update(msg, out changed, out added);
                            if (added)
                            {
                                OnNewFax(new FaxEventArgs(msg));
                            }
                            else
                            {
                                OnFaxUpdated(new FaxEventArgs(msg));
                            }
                        }
                        else
                        {
                            _FaxCache.Add(msg);
                            OnNewFax(new FaxEventArgs(msg));
                        }
                        messages.Add(msg);
                    }

                    // Now find any messages in the cache that
                    // no longer exist in the XML data sent from the
                    // server and mark them as deleted
                    FaxMessage[] deletedMessages = _FaxCache.RemoveMessagesExcept(messages.ToArray());
                    if ((deletedMessages != null) && (deletedMessages.Length > 0))
                    {
                        foreach (FaxMessage msg in deletedMessages)
                        {
                            OnFaxDeleted(new FaxEventArgs(msg));
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        #endregion

        #endregion

        #region Voicemail

        /// <summary>
        /// Gets the existence of waiting (unread) voicemails.
        /// </summary>
        /// <remarks>
        /// <ininWatchRequired />
        /// </remarks>
        /// <value><see langword="true"/> if a voicemail is waiting; otherwise, <see langword="false"/>.</value>
        /// <exception cref="ININ.IceLib.NotCachedException">VoicemailWaiting not being watched.</exception>
        public bool VoicemailWaiting
        {
            get
            {
                using (TraceContext.SessionIdFromSession(Session))
                {
                    try
                    {
                        if (_WatchingVMWaitingRefCount > 0)
                        {
                            return _VoicemailWaiting;
                        }

                            throw new NotCachedException("VoicemailWaiting");
                        }
                    catch (Exception ex)
                    {
                        // Record the exception and rethrow.
                        TraceTopics.UnifiedMessaging.exception(ex);
                        throw;
                    }
                }
            }
        }

        #region Voicemail Events

        /// <summary>
        /// Occurs when the VoicemailWaiting value changes.
        /// </summary>
        /// <remarks>
        /// This event will only occur if voicemails are being watched. To start watching
        /// voicemails either the <see cref="StartWatchingVoicemailWaiting"/> method or
        /// the <see cref="StartWatchingVoicemailWaitingAsync"/> method.
        /// <note>All event handlers should be added before calling
        /// <see cref="StartWatchingVoicemailWaiting"/> or 
        /// <see cref="StartWatchingVoicemailWaitingAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// </remarks>
        public event EventHandler VoicemailWaitingChanged;

        /// <summary>
        /// Occurs when a new voicemail is received.
        /// </summary>
        /// <remarks>
        /// This event will only occur if voicemails are being watched. To start watching
        /// voicemails either the <see cref="StartWatchingVoicemailWaiting"/> method or
        /// the <see cref="StartWatchingVoicemailWaitingAsync"/> method.
        /// <note>All event handlers should be added before calling
        /// <see cref="StartWatchingVoicemailWaiting"/> or 
        /// <see cref="StartWatchingVoicemailWaitingAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// </remarks>
        public event EventHandler<VoicemailEventArgs> NewVoicemail;

        /// <summary>
        /// Occurs when a voicemail is deleted.
        /// </summary>
        /// <remarks>
        /// This event will only occur if voicemails are being watched. To start watching
        /// voicemails either the <see cref="StartWatchingVoicemailWaiting"/> method or
        /// the <see cref="StartWatchingVoicemailWaitingAsync"/> method.
        /// <note>All event handlers should be added before calling
        /// <see cref="StartWatchingVoicemailWaiting"/> or 
        /// <see cref="StartWatchingVoicemailWaitingAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// </remarks>
        public event EventHandler<VoicemailEventArgs> VoicemailDeleted;

        /// <summary>
        /// Occurs when a voicemail is modified.
        /// </summary>
        /// <remarks>
        /// This event will only occur if voicemails are being watched. To start watching
        /// voicemails either the <see cref="StartWatchingVoicemailWaiting"/> method or
        /// the <see cref="StartWatchingVoicemailWaitingAsync"/> method.
        /// <note>All event handlers should be added before calling
        /// <see cref="StartWatchingVoicemailWaiting"/> or 
        /// <see cref="StartWatchingVoicemailWaitingAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// </remarks>
        public event EventHandler<VoicemailEventArgs> VoicemailUpdated;

        /// <summary>
        /// Occurs when an operation begins updating the list of voicemails.
        /// </summary>
        /// <remarks><para>This event occurs when an operation, such as <see cref="ININ.IceLib.UnifiedMessaging.UnifiedMessagingManager.RefreshVoicemailCache"/>, 
        /// begins updating the local list of voicemails. A corresponding <see cref="ININ.IceLib.UnifiedMessaging.UnifiedMessagingManager.VoicemailUpdateCompleted"/> event will
        /// will occur when the operation completes.</para>
        /// This event will only occur if voicemails are being watched. To start watching
        /// voicemails either the <see cref="StartWatchingVoicemailWaiting"/> method or
        /// the <see cref="StartWatchingVoicemailWaitingAsync"/> method.
        /// <note>All event handlers should be added before calling
        /// <see cref="StartWatchingVoicemailWaiting"/> or 
        /// <see cref="StartWatchingVoicemailWaitingAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// </remarks>
        public event EventHandler<EventArgs> VoicemailUpdateStarted;

        /// <summary>
        /// Occurs when an operation finishes updating the list of voicemails.
        /// </summary>
        /// <remarks><para>This event occurs when an operation, such as <see cref="ININ.IceLib.UnifiedMessaging.UnifiedMessagingManager.RefreshVoicemailCache"/>, 
        /// finishes updating the local list of voicemails. A corresponding <see cref="ININ.IceLib.UnifiedMessaging.UnifiedMessagingManager.VoicemailUpdateStarted"/> event
        /// occurs when the operation begins.</para>
        /// This event will only occur if voicemails are being watched. To start watching
        /// voicemails either the <see cref="StartWatchingVoicemailWaiting"/> method or
        /// the <see cref="StartWatchingVoicemailWaitingAsync"/> method.
        /// <note>All event handlers should be added before calling
        /// <see cref="StartWatchingVoicemailWaiting"/> or 
        /// <see cref="StartWatchingVoicemailWaitingAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// </remarks>
        public event EventHandler<EventArgs> VoicemailUpdateCompleted;

        //This mutex is locked while processing voicemail list updates from Session Manager.
        private readonly string _VoicemailsUpdatingMutex = String.Empty;

        /// <summary>
        /// Occurs when the server has completed playing a voicemail to the user.
        /// </summary>
        /// <remarks>
        /// This event will only occur if voicemails are being watched. To start watching
        /// voicemails either the <see cref="StartWatchingVoicemailWaiting"/> method or
        /// the <see cref="StartWatchingVoicemailWaitingAsync"/> method.
        /// <note>All event handlers should be added before calling
        /// <see cref="StartWatchingVoicemailWaiting"/> or 
        /// <see cref="StartWatchingVoicemailWaitingAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// <para>This event will only occur when <see cref="VoicemailMessage.PlayToHandset"/> or 
        /// <see cref="VoicemailMessage.PlayToHandsetAsync"/> is invoked.</para>
        /// </remarks>
        /// <seealso cref="VoicemailMessage.PlayToHandset"/>
        /// <seealso cref="VoicemailMessage.PlayToHandsetAsync"/>
        public event EventHandler<VoicemailServerPlayResultEventArgs> VoicemailServerPlayResult;

        private void OnVoicemailWaitingChanged(EventArgs e)
        {
            EventHelper.SafeRaise(VoicemailWaitingChanged, this, e, TraceTopics.UnifiedMessaging);
        }

        private void OnNewVoicemail(VoicemailEventArgs e)
        {
            EventHelper.SafeRaise(NewVoicemail, this, e, TraceTopics.UnifiedMessaging);
        }

        private void OnVoicemailDeleted(VoicemailEventArgs e)
        {
            EventHelper.SafeRaise(VoicemailDeleted, this, e, TraceTopics.UnifiedMessaging);
        }

        private void OnVoicemailUpdated(VoicemailEventArgs e)
        {
            EventHelper.SafeRaise(VoicemailUpdated, this, e, TraceTopics.UnifiedMessaging);
        }

        private void OnVoiceMessageServerPlayResult(VoicemailServerPlayResultEventArgs e)
        {
            EventHelper.SafeRaise(VoicemailServerPlayResult, this, e, TraceTopics.UnifiedMessaging);
        }

        #endregion

        #region StartWatchingVoicemailWaiting

        private void StartWatchingVoicemailWaitingImpl()
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    _Session.CheckSession();

                    Message request = _Session.CreateMessage(InternalThinManagerEventIds.eThinManagerReq_WatchMessageLightEvents);
                    request.Payload.Writer.Write(true);

                    _Session.SendMessageAndThrowIfError(request);

                    _WatchingVMWaitingRefCount++;
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues a synchronous request to watch for waiting voicemails.
        /// </summary>
        /// <remarks>
        /// <note>All event handlers should be added before calling
        /// <see cref="StartWatchingVoicemailWaiting"/> or 
        /// <see cref="StartWatchingVoicemailWaitingAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// </remarks>
        /// <ConnectionExceptions />
        public void StartWatchingVoicemailWaiting()
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    StartWatchingVoicemailWaitingImpl();
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues an asynchronous request to watch for waiting voicemails.
        /// </summary>
        /// <param name="completedCallback">The callback to invoke when the asynchronous operation completes.</param>
        /// <param name="userState"><ininAsyncStateParam /></param>
        /// <remarks>
        /// <ininAsyncMethodNote />
        /// <note>All event handlers should be added before calling
        /// <see cref="StartWatchingVoicemailWaiting"/> or 
        /// <see cref="StartWatchingVoicemailWaitingAsync"/>.</note>
        /// <ininHowWatchesWork />
        /// </remarks>
        public void StartWatchingVoicemailWaitingAsync(AsyncCompletedEventHandler completedCallback, object userState)
        {
            long taskId = _AsyncTaskTracker.CreateTaskId();
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(taskId))
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    AsyncUnifiedMessagingManagerState asyncState = new AsyncUnifiedMessagingManagerState(taskId, completedCallback, userState);
                    _AsyncTaskTracker.PerformTask(asyncState, StartWatchingVoicemailWaitingAsyncPerformTask, StartWatchingVoicemailWaitingAsyncCompleted);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async operation (called on an arbitrary worker thread).
        private void StartWatchingVoicemailWaitingAsyncPerformTask(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    StartWatchingVoicemailWaitingImpl();
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async completion (called on the UI thread).
        private void StartWatchingVoicemailWaitingAsyncCompleted(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            {
                AsyncCompletedEventArgs e = new AsyncCompletedEventArgs(
                        asyncState.Exception,
                        asyncState.Cancelled,
                        asyncState.UserState);
                EventHelper.SafeRaise(asyncState.CompletedCallback, this, e, TraceTopics.UnifiedMessaging);
            }
        }

        #endregion // StartWatchingVoicemailWaiting

        #region StopWatchingVoicemailWaiting

        private void StopWatchingVoicemailWaitingImpl()
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                if (_WatchingVMWaitingRefCount > 0)
                {
                    // Decrement refcount and check
                    if (--_WatchingVMWaitingRefCount == 0)
                    {
                        if (_Session.ConnectionState != ConnectionState.Up)
                        {
                            TraceTopics.UnifiedMessaging.status("Session.ConnectionState is not Up. Session is in the disconnected state");
                            return; // we do not need to send a message if the session is disconnected.
                        }

                        try
                        {

                            Message request = _Session.CreateMessage(InternalThinManagerEventIds.eThinManagerReq_WatchMessageLightEvents);
                            request.Payload.Writer.Write(false);
                            _Session.SendMessageAndThrowIfError(request);
                        }
                        catch (Exception ex)
                        {
                            TraceTopics.UnifiedMessaging.exception(ex, "Error stopping watch on VoicemailWaiting");
                            throw;
                        }
                    }
                }
                else
                {
                    TraceTopics.UnifiedMessaging.warning("No watchers are remaining for VoicemailWaiting, no need to stop watching (ignoring request)");
                }
            }
        }

        /// <summary>
        /// Issues a synchronous request to stop watching for waiting voicemails.
        /// </summary>
        /// <remarks>
        /// <ininHowWatchesWork />
        /// </remarks>
        /// <ConnectionExceptions />
        public void StopWatchingVoicemailWaiting()
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    StopWatchingVoicemailWaitingImpl();
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues an asynchronous request to stop watching for waiting voicemails.
        /// </summary>
        /// <param name="completedCallback">The callback to invoke when the asynchronous operation completes.</param>
        /// <param name="userState"><ininAsyncStateParam /></param>
        /// <remarks>
        /// <ininAsyncMethodNote />
        /// <ininHowWatchesWork />
        /// </remarks>
        public void StopWatchingVoicemailWaitingAsync(AsyncCompletedEventHandler completedCallback, object userState)
        {
            long taskId = _AsyncTaskTracker.CreateTaskId();
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(taskId))
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    AsyncUnifiedMessagingManagerState asyncState = new AsyncUnifiedMessagingManagerState(taskId, completedCallback, userState);
                    _AsyncTaskTracker.PerformTask(asyncState, StopWatchingVoicemailWaitingAsyncPerformTask, StopWatchingVoicemailWaitingAsyncCompleted);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async operation (called on an arbitrary worker thread).
        private void StopWatchingVoicemailWaitingAsyncPerformTask(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    StopWatchingVoicemailWaitingImpl();
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async completion (called on the UI thread).
        private void StopWatchingVoicemailWaitingAsyncCompleted(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            {
                AsyncCompletedEventArgs e = new AsyncCompletedEventArgs(
                        asyncState.Exception,
                        asyncState.Cancelled,
                        asyncState.UserState);
                EventHelper.SafeRaise(asyncState.CompletedCallback, this, e, TraceTopics.UnifiedMessaging);
            }
        }

        #endregion // StopWatchingVoicemailWaiting

        #region StopVoicemailHandsetPlayback

        private void StopVoicemailHandsetPlaybackImpl()
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    Message request = _Session.CreateMessage(InternalThinManagerEventIds.eThinManagerReq_VoiceMessageStop);
                    _Session.SendMessageAndThrowIfError(request);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues a synchronous request to stop voicemail playback on the handset.
        /// </summary>
        /// <ConnectionExceptions />
        public void StopVoicemailHandsetPlayback()
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    StopVoicemailHandsetPlaybackImpl();
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues an asychronous request to stop voicemail playback on the handset.
        /// </summary>
        /// <param name="completedCallback">The callback to invoke when the asynchronous operation completes.</param>
        /// <param name="userState"><ininAsyncStateParam /></param>
        /// <remarks>
        /// <ininAsyncMethodNote />
        /// </remarks>
        public void StopVoicemailHandsetPlaybackAsync(AsyncCompletedEventHandler completedCallback, object userState)
        {
            long taskId = _AsyncTaskTracker.CreateTaskId();
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(taskId))
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    AsyncUnifiedMessagingManagerState asyncState = new AsyncUnifiedMessagingManagerState(taskId, completedCallback, userState);
                    _AsyncTaskTracker.PerformTask(asyncState, StopVoicemailHandsetPlaybackAsyncPerformTask, StopVoicemailHandsetPlaybackAsyncCompleted);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async operation (called on an arbitrary worker thread).
        private void StopVoicemailHandsetPlaybackAsyncPerformTask(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    StopVoicemailHandsetPlaybackImpl();
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async completion (called on the UI thread).
        private void StopVoicemailHandsetPlaybackAsyncCompleted(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            {
                AsyncCompletedEventArgs e = new AsyncCompletedEventArgs(
                        asyncState.Exception,
                        asyncState.Cancelled,
                        asyncState.UserState);
                EventHelper.SafeRaise(asyncState.CompletedCallback, this, e, TraceTopics.UnifiedMessaging);
            }
        }

        #endregion // StopVoicemailHandsetPlayback

        #region RefreshVoicemailCache

        private void RefreshVoicemailCacheImpl(int maxMessages)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked. maxMessages={}", maxMessages);

                try
                {
                    try
                    {
                        if (maxMessages < -1)
                        {
                            maxMessages = -1;
                        }

                        Message request = _Session.CreateMessage(InternalThinManagerEventIds.eThinManagerReq_GetVoiceMessageList);
                        request.Payload.Writer.Write(maxMessages);
                        _Session.SendMessageAndThrowIfError(request);
                    }
                    catch (Exception ex)
                    {
                        if (((IceLibException)ex).ErrorCode == InternalHRESULTs.E_I3ERR_REQUEST_REJECTED)
                        {
                            throw new RateLimitedException(ex);
                        }

                        throw new IceLibException(String.Format(Localization.LoadString(_Session, "REFRESH_LIST_ERROR"), "Voicemail"), ex);
                    }
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues a synchronous request to refresh the voicemail cache.
        /// </summary>
        /// <remarks>
        /// Any subsequent calls to this method or its async version (<see cref="RefreshVoicemailCacheAsync"/>) within a 5 second period will be rejected with a <see cref="RateLimitedException"/>.
        /// This restriction is per message type, a call to <see cref="RefreshVoicemailCache"/> followed by a call to <see cref="RefreshFaxCache"/> or <see cref="RefreshFaxCacheAsync"/> within a 5 second period will be successful.
        /// </remarks>
        /// <param name="maxMessages">The number of messages to refresh, -1 to refresh them all.</param>
        /// <exception cref="IceLibException">There was an error refreshing the list or this method was called more than once within a 5 second period.</exception>
        /// <ConnectionExceptions />
        public void RefreshVoicemailCache(int maxMessages)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked. maxMessages={}", maxMessages);

                try
                {
                    RefreshVoicemailCacheImpl(maxMessages);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues an asynchronous request to refresh the voicemail cache.
        /// </summary>
        /// <remarks>
        /// Any subsequent calls to this method or its synchronous version (<see cref="RefreshVoicemailCache"/>) within a 5 second period will be rejected with a <see cref="RateLimitedException"/>.
        /// This restriction is per message type, a call to <see cref="RefreshVoicemailCacheAsync"/> followed by a call to <see cref="RefreshFaxCache"/> or <see cref="RefreshFaxCacheAsync"/> within a 5 second period will be successful.
        /// </remarks>
        /// <param name="maxMessages">The number of messages to refresh, -1 to refresh them all.</param>
        /// <param name="completedCallback">The callback to invoke when the asynchronous operation completes.</param>
        /// <param name="userState"><ininAsyncStateParam /></param>
        /// <remarks><ininAsyncMethodNote /></remarks>
        public void RefreshVoicemailCacheAsync(int maxMessages, AsyncCompletedEventHandler completedCallback, object userState)
        {
            long taskId = _AsyncTaskTracker.CreateTaskId();
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(taskId))
            {
                TraceTopics.UnifiedMessaging.status("Invoked. maxMessages={}", maxMessages);

                try
                {
                    AsyncUnifiedMessagingManagerState asyncState = new AsyncUnifiedMessagingManagerState(maxMessages, taskId, completedCallback, userState);
                    _AsyncTaskTracker.PerformTask(asyncState, RefreshVoicemailCacheAsyncPerformTask, RefreshVoicemailCacheAsyncCompleted);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async operation (called on an arbitrary worker thread).
        private void RefreshVoicemailCacheAsyncPerformTask(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked. maxMessages={}", asyncState.MaxMessages);

                try
                {
                    RefreshVoicemailCacheImpl(asyncState.MaxMessages);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async completion (called on the UI thread).
        private void RefreshVoicemailCacheAsyncCompleted(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            {
                AsyncCompletedEventArgs e = new AsyncCompletedEventArgs(
                        asyncState.Exception,
                        asyncState.Cancelled,
                        asyncState.UserState);
                EventHelper.SafeRaise(asyncState.CompletedCallback, this, e, TraceTopics.UnifiedMessaging, asyncState.MaxMessages);
            }
        }

        #endregion // RefreshVoicemailCache

        #region UpdateMessageWaitingIndicator

        private void UpdateMessageWaitingIndicatorImpl(int actualMessageCount)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked. messageCount={}", actualMessageCount);

                try
                {
                    Message request = _Session.CreateMessage(InternalThinManagerEventIds.eThinManagerReq_ProcessMessageLight);
                    request.Payload.Writer.Write(actualMessageCount);
                    _Session.SendMessageAndThrowIfError(request);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues a synchronous request to update the message waiting indicator.
        /// </summary>
        /// <param name="actualMessageCount">The count of messages, or -1 to force the server to compute the count.</param>
        /// <ConnectionExceptions />
        public void UpdateMessageWaitingIndicator(int actualMessageCount)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    UpdateMessageWaitingIndicatorImpl(actualMessageCount);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Issues an asynchronous request to update the message waiting indicator.
        /// </summary>
        /// <param name="actualMessageCount">The count of messages, or -1 to force the server to compute the count.</param>
        /// <param name="completedCallback">The callback to invoke when the asynchronous operation completes.</param>
        /// <param name="userState"><ininAsyncStateParam /></param>
        /// <remarks><ininAsyncMethodNote /></remarks>
        public void UpdateMessageWaitingIndicatorAsync(int actualMessageCount, AsyncCompletedEventHandler completedCallback, object userState)
        {
            long taskId = _AsyncTaskTracker.CreateTaskId();
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(taskId))
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    AsyncUnifiedMessagingManagerState asyncState = new AsyncUnifiedMessagingManagerState(actualMessageCount, taskId, completedCallback, userState);
                    _AsyncTaskTracker.PerformTask(asyncState, UpdateMessageWaitingIndicatorAsyncPerformTask, UpdateMessageWaitingIndicatorAsyncCompleted);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async operation (called on an arbitrary worker thread).
        private void UpdateMessageWaitingIndicatorAsyncPerformTask(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            using (TraceTopics.UnifiedMessaging.scope())
            {
                TraceTopics.UnifiedMessaging.status("Invoked.");

                try
                {
                    UpdateMessageWaitingIndicatorImpl(asyncState.ActualMessageCount);
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        // This callback performs the actual async completion (called on the UI thread).
        private void UpdateMessageWaitingIndicatorAsyncCompleted(AsyncUnifiedMessagingManagerState asyncState)
        {
            using (TraceContext.SessionIdFromSession(_Session))
            using (TraceContext.IceLibTaskId.create(asyncState.TaskId))
            {
                AsyncCompletedEventArgs e = new AsyncCompletedEventArgs(
                        asyncState.Exception,
                        asyncState.Cancelled,
                        asyncState.UserState);
                EventHelper.SafeRaise(asyncState.CompletedCallback, this, e, TraceTopics.UnifiedMessaging);
            }
        }

        #endregion // UpdateMessageWaitingIndicator

        #region Voicemail Methods

        /// <summary>
        /// Gets a <see cref="System.Collections.ObjectModel.ReadOnlyCollection{T}"/> of
        /// all VoicemailMessages.
        /// </summary>
        /// <remarks>
        /// <note>The Voicemails property will not be populated until after either the 
        /// <see cref="RefreshVoicemailCache"/> method or the 
        /// <see cref="RefreshVoicemailCacheAsync" /> method has been called and the 
        /// <see cref="NewVoicemail"/> event has fired.</note>
        /// </remarks>
        /// <value>The voicemail messages.</value>
        public ReadOnlyCollection<VoicemailMessage> Voicemails
        {
            get { return Array.AsReadOnly<VoicemailMessage>(_VMCache.GetAll()); }
        }

        internal void RemoveFromCache(VoicemailMessage message)
        {
            _VMCache.Remove(message);
            OnVoicemailDeleted(new VoicemailEventArgs(message));
        }

        #endregion

        #region Voicemail Message Handlers

        private void HandleMessageLight(object sender, MessageReceivedEventArgs e)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                try
                {
                    bool bOn = e.Message.Payload.Reader.ReadBoolean();
                    if (_VoicemailWaiting != bOn)
                    {
                        _VoicemailWaiting = bOn;
                        OnVoicemailWaitingChanged(EventArgs.Empty);
                    }
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        private void HandleVoiceMessageList(object sender, MessageReceivedEventArgs e)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                try
                {
                    lock (_VoicemailsUpdatingMutex)
                    {
                        EventHelper.SafeRaise(VoicemailUpdateStarted, this, EventArgs.Empty, TraceTopics.UnifiedMessaging);

                        string sXML = e.Message.Payload.Reader.ReadString();
                        StringReader sr = new StringReader(sXML);
                        XmlVoicemailList list = (XmlVoicemailList) _XmlVoicemailListSerializer.Deserialize(sr);
                        List<VoicemailMessage> messages = new List<VoicemailMessage>();

                        // Start by adding all items to the collection.  For any
                        // that don't already exist we send a notification, if an existing voicemail
                        // has changed, send the changed notification
                        foreach (XmlMessage message in list.AllMessages)
                        {
                            VoicemailMessage msg = new VoicemailMessage(message, this);

                            if (_VMCache.Contains(msg))
                            {
                                bool changed;
                                bool added;

                                _VMCache.Update(msg, out changed, out added);

                                if (added)
                                {
                                    OnNewVoicemail(new VoicemailEventArgs(msg));
                                }
                                else
                                {
                                    OnVoicemailUpdated(new VoicemailEventArgs(msg));
                                }
                            }
                            else
                            {
                                _VMCache.Add(msg);
                                OnNewVoicemail(new VoicemailEventArgs(msg));
                            }

                            messages.Add(msg);
                        }

                        // Now find any messages in the cache that
                        // no longer exist in the XML data sent from the
                        // server and mark them as deleted
                        VoicemailMessage[] deletedMessages = _VMCache.RemoveMessagesExcept(messages.ToArray());

                        if ((deletedMessages != null) && (deletedMessages.Length > 0))
                        {
                            foreach (VoicemailMessage msg in deletedMessages)
                            {
                                OnVoicemailDeleted(new VoicemailEventArgs(msg));
                            }
                        }

                        EventHelper.SafeRaise(VoicemailUpdateCompleted, this, EventArgs.Empty, TraceTopics.UnifiedMessaging);
                    } //unlock(_VoicemailsUpdatingMutex)
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        private void HandleMessagePlayResult(object sender, MessageReceivedEventArgs e)
        {
            using (TraceTopics.UnifiedMessaging.scope())
            {
                try
                {
                    // Update to use internal.ReadLastError
                    int resultCode;
                    string resultText;
                   
                    StreamErrorInformation.ReadLastError(e.Message.Payload.Reader, out resultCode, out resultText, true);

                    VoicemailMessage message;
                    VoicemailAttachment attachment = null;
                    string messageId;

                    try
                    {
                        messageId = e.Message.Payload.Reader.ReadString();

                        string attachmentId = e.Message.Payload.Reader.ReadString();

                        // Find message
                        message = _VMCache.Get(messageId);

                        if ((message != null) && (message.Attachments != null))
                        {
                            // Find Attachment
                            for (int i = 0; i < message.Attachments.Count; i++)
                            {
                                if (String.Compare(message.Attachments[i].Id, attachmentId) == 0)
                                {
                                    attachment = message.Attachments[i];
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // The structure has no message or attachment Id provided.
                        // It's most likely from an older SM
                        TraceTopics.UnifiedMessaging.exception(ex, "Unable to read Message or Attachment ID, possible connection to old SM.");
                        throw;
                    }

                    // Determine if the VoiceMessageServerPlayResult event or VoicemailMessage.PlaybackComplete event.
                    VoicemailPlayEventTarget voicemailPlayEventTarget;

                    lock (_VoicemailPlayEventTargets)
                    {
                        // Retrieve the value if it exists, and then clear it out of cache.
                        _VoicemailPlayEventTargets.TryGetValue(messageId, out voicemailPlayEventTarget);
                        _VoicemailPlayEventTargets.Remove(messageId);
                    }

                    switch (voicemailPlayEventTarget)
                    {
                        case VoicemailPlayEventTarget.VoiceMessageServerPlayResult:
                            TraceTopics.UnifiedMessaging.verbose(
                                "Invoking VoicemailServerPlayResult since the voicemail play event target was '{}'.",
                                voicemailPlayEventTarget);

                            var result = new VoiceMessageServerPlayResult(message, attachment, resultCode, resultText);
                            OnVoiceMessageServerPlayResult(new VoicemailServerPlayResultEventArgs(result));

                            break;
                        case VoicemailPlayEventTarget.VoicemailMessagePlaybackComplete:
                            TraceTopics.UnifiedMessaging.verbose(
                                    "Invoking VoicemailMessage.PlaybackComplete since the voicemail play event target was '{}'.",
                                    voicemailPlayEventTarget);

                            if (message != null)
                            {
                                message.OnPlayWaveAudioComplete(this, EventArgs.Empty);
                            }

                            break;
                        default:
                            TraceTopics.UnifiedMessaging.verbose("Voicemail play event target was '{}', no events to invoke.",
                                                                 voicemailPlayEventTarget);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    // Record the exception and rethrow.
                    TraceTopics.UnifiedMessaging.exception(ex);
                    throw;
                }
            }
        }

        #endregion

        internal void SetVoicemailPlayEventTarget(string messageId, VoicemailPlayEventTarget voicemailPlayEventTarget)
        {
            lock (_VoicemailPlayEventTargets)
            {
                _VoicemailPlayEventTargets[messageId] = voicemailPlayEventTarget;
            }
        }

        internal void ClearVoicemailPlayEventTarget(string messageId)
        {
            lock (_VoicemailPlayEventTargets)
            {
                _VoicemailPlayEventTargets.Remove(messageId);
            }
        }

        #endregion

        // Example URI
        // imap://cpallott@voicestore:993/INBOX;uidvalidity=1282171255/;uid=574/;section=1
        internal static string GetUidFromUri(string id)
        {
            var  retVal = id;
            var regex = new Regex(_UidRegex, RegexOptions.IgnoreCase);
            var match = regex.Match(id);

            if (match.Success && match.Groups[_UidRegexUidGroup] != null)
            {
                retVal = match.Groups[_UidRegexUidGroup].Value;
            }

            return retVal;
        }

        /// <summary>
        /// Handler for XML parser when an Unknown Node is encountered.
        /// </summary>
        /// <param name="sender">The object sending the event message.</param>
        /// <param name="e">XML Node data for the Unknown Node event.</param>
        private static void XML_UnknownNode(object sender, XmlNodeEventArgs e)
        {
            TraceTopics.UnifiedMessaging.error("XML_UnknownNode Name={}, Text={}", e.Name, e.Text);
        }

        /// <summary>
        /// Handler for XML parser when an Unknown Attribute is encountered.
        /// </summary>
        /// <param name="sender">The object sending the event message.</param>
        /// <param name="e">XML Node data for the Unknown Attribute event.</param>
        private static void XML_UnknownAttribute(object sender, XmlAttributeEventArgs e)
        {
            TraceTopics.UnifiedMessaging.error("XML_UnknownAttribute Name={}, Value={}", e.Attr.Name, e.Attr.Value);
        }
    }
}
