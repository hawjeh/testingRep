using System;
using Microsoft.DirectX.DirectSound;
using System.Diagnostics;

namespace ININ.Audio
{
	/// <summary>
	/// The play state of the device.
	/// </summary>
	public enum PlayState
	{
		/// <summary>
		/// The Playing PlayState.
		/// </summary>
		Playing, 

		/// <summary>
		/// The Stopped PlayState.
		/// </summary>
		Stopped, 

		/// <summary>
		/// The Paused PlayState.
		/// </summary>
		Paused
	};

	/// <summary>
	/// Wraps a DirectX Managed Audio file. The IAudio interface was based
	/// on the DirectX Audio class, so it's a one-to-one mapping.
	/// </summary>
	internal class ManagedDirectSoundAudio : IAudio, IDisposable
	{
		private Device					_Device = null;
		private StreamingSoundBuffer	_Buffer = null;
		private bool					m_Disposed;

        /// <summary>
        /// Creates a Device object which will loads and initializes the DirectSound DLL. 
        /// By creating the device without specifying an owner, hoping to get the inital 
        /// load out of the way withoutrunning into any LoaderLock issues within DirectSound.
        /// </summary>
        public static void InitAudioEngine()
        {
            try
            {
                Device d = new Device(Guid.Empty);
                d.Dispose();
                d = null;
            }
            catch (Exception ex)
            {
                string error = String.Format("DirectSoundAudio.InitAudioEngine Exception: {0}", ex.Message);
                Debug.WriteLine(error);
            }
        }

		/// <summary>
		/// Creates a DirectSoundAudio object that plays from the specified filename to the specified Device.
		/// </summary>
		/// <param name="owner">The owner of this audio (required to handle application focus switching)</param>
		/// <param name="deviceId">The Guid of the device to play to.</param>
		/// <param name="filename">The file containing the audio data.</param>
		public ManagedDirectSoundAudio(System.Windows.Forms.Control owner, Guid deviceId, string filename) : this(owner, deviceId)
		{
			Open(filename);
		}
		
		/// <summary>
		/// Creates a DirectSoundAudio object that plays from the specified filename to the specified Device.
		/// </summary>
		/// <param name="owner">The owner of this audio (required to handle application focus switching)</param>
		/// <param name="deviceId">The Guid of the device to play to.</param>
		/// <param name="filename">The file containing the audio data.</param>
		public ManagedDirectSoundAudio(IntPtr owner, Guid deviceId, string filename) : this(owner, deviceId)
		{
			Open(filename);
		}

		/// <summary>
		/// Ctor.
		/// </summary>
		protected ManagedDirectSoundAudio(System.Windows.Forms.Control owner, Guid deviceId)
		{
			if (deviceId == Guid.Empty)
			{
				_Device = new Device();
			}
			else
			{
				_Device = new Device(deviceId);
			}
			_Device.SetCooperativeLevel(owner, CooperativeLevel.Priority);
		}
		
		/// <exclude />
		protected ManagedDirectSoundAudio(IntPtr owner, Guid deviceId)
		{
			if (deviceId == Guid.Empty)
			{
				_Device = new Device();
			}
			else
			{
				_Device = new Device(deviceId);
			}
			_Device.SetCooperativeLevel(owner, CooperativeLevel.Priority);
		}

		private void Open(string filename)
		{
			_Buffer = new StreamingSoundBuffer(filename, _Device);
			_Buffer.Starting += new EventHandler(_Buffer_Starting);
			_Buffer.Stopping += new EventHandler(_Buffer_Stopping);
			_Buffer.Pausing += new EventHandler(_Buffer_Pausing);
		}

		#region IAudio Members

        private void AssertNonNullBuffer()
        {
            if (_Buffer == null)
            {
                throw new NullReferenceException(Localization.LoadString("ERROR_NO_AUDIO"));
            }
        }

		/// <summary>
		/// The current position within the audio.
		/// </summary>
		/// <remarks>This value's valid range is 0 - [data length] 
		/// where data length is the length of the original audio file in bytes.</remarks>
		public long CurrentPosition
		{
			get
			{
                AssertNonNullBuffer();
				return _Buffer.CurrentPosition;
			}
			set
			{
                AssertNonNullBuffer();
				_Buffer.CurrentPosition = value;
			}
		}

		/// <summary>
		/// The length of the available audio.
		/// </summary>
		/// <remarks>This value is equal to the number of bytes representing the audio
		/// in the original wave file.</remarks>
		public long Duration
		{
			get
			{
                AssertNonNullBuffer();
				return _Buffer.Duration;
			}
		}

		/// <summary>
		/// The number of bytes that represent one second for this wave file.
		/// </summary>
		/// <remarks>This is used in conjunction with Duration and/or CurrentPosition to translate
		/// values from bytes to seconds.</remarks>
		public long BytesPerSecond
		{
			get
			{
                AssertNonNullBuffer();
				return _Buffer.BytesPerSecond;
			}
		}

		private const int MinVolume = -3500;

		/// <summary>
		/// The volume of the audio playback.
		/// </summary>
		/// <remarks>The Volume value ranges from 0 (muted) to 100 (full volume).</remarks>
		public int Volume
		{
			get
			{
                AssertNonNullBuffer();

				if (_Buffer.Volume == (int)Microsoft.DirectX.DirectSound.Volume.Min)
				{
					// 0 = mute = volume.min
					return 0;
				}
				else if (_Buffer.Volume == 0)
				{
					return 100;
				}
				else
				{
					// Scale -3500 to 0 (decibels) back to 1 - 100
					return Convert.ToInt32(Math.Floor((100 - ((double)_Buffer.Volume / (double)MinVolume * 100))));
				}
			}
			set
			{
                if ((value < 0) || (value > 100))
                {
                    throw new ArgumentOutOfRangeException("Volume", "Volume must be between 0 and 100");
                }

                AssertNonNullBuffer();
				
				if (value == 0)
				{
					_Buffer.Volume = (int)Microsoft.DirectX.DirectSound.Volume.Min;
				}
				else if (value == 100)
				{
					_Buffer.Volume = 0;
				}
				else
				{
					// Since the audio becomes essentially muted close to -4000 decibels,
					// scale the input value between -3500 and -1
					_Buffer.Volume = (MinVolume * (100 - value)) / 100;
				}
			}
		}

		/// <summary>
		/// -100 (left) to 0 (center) to 100 (right)
		/// </summary>
		public int Pan
		{
			get
			{
                AssertNonNullBuffer();
				return _Buffer.Pan;
			}
			set
			{
                if ((value < -100) || (value > 100))
                {
                    throw new ArgumentOutOfRangeException("Pan", String.Format("Pan must be between {0} and {1}", (int)Microsoft.DirectX.DirectSound.Pan.Left, (int)Microsoft.DirectX.DirectSound.Pan.Right));
                }

                AssertNonNullBuffer();
				// Scale to the normal range of -10000 to 10000
				_Buffer.Pan = value * 100;
			}
		}

		/// <summary>
		/// Valid values are -100 to 100 (where -100 is half the speed, and 100 is double the speed)
		/// </summary>
		public int Frequency
		{
			get
			{
                AssertNonNullBuffer();

				if (_Buffer.Frequency == _Buffer.DefaultFrequency)
				{
					return 0;
				}
				else
				{
					double adjust = (_Buffer.Frequency > _Buffer.DefaultFrequency) ? 100 : 200;
					return Convert.ToInt32(((double)(_Buffer.Frequency - _Buffer.DefaultFrequency) * adjust) / (double)_Buffer.DefaultFrequency);
				}
			}
			set
			{
                if (((value < -100) || (value > 100)) && (value != 0))
                {
                    throw new ArgumentOutOfRangeException("Frequency", "Frequency must be between -100 and 100");
                }


                AssertNonNullBuffer();

				int newFreq = Convert.ToInt32(_Buffer.DefaultFrequency);
				if (value != 0)
				{
					double adjust = (value > 0) ? 100 : 200;
					newFreq += Convert.ToInt32((double)newFreq * ((double)value / adjust));
				}
				_Buffer.Frequency = newFreq;
			}
		}

		/// <summary>
		/// Gets whether the audio is paused.
		/// </summary>
		public bool Paused
		{
			get
			{
                AssertNonNullBuffer();
				return _Buffer.Paused;
			}
		}

		/// <summary>
		/// Gets whether the audio is playing.
		/// </summary>
		public bool Playing
		{
			get
			{
                AssertNonNullBuffer();
				return _Buffer.Playing;
			}
		}

		/// <summary>
		/// Gets whether the audio is stopped.
		/// </summary>
		public bool Stopped
		{
			get
			{
                AssertNonNullBuffer();
				return _Buffer.Stopped;
			}
		}

		/// <summary>
		/// Occurs when the audio starts.
		/// </summary>
		public event System.EventHandler Starting;

		/// <summary>
		/// Occurs when the audio pauses.
		/// </summary>
		public event System.EventHandler Pausing;

		/// <summary>
		/// Occurs when the audio stops.
		/// </summary>
		public event System.EventHandler Stopping;

		/// <summary>
		/// Pauses the audio playback.
		/// </summary>
		/// <returns></returns>
		public long Pause()
		{
            AssertNonNullBuffer();

			return _Buffer.Pause();
		}

		/// <summary>
		/// Begins playing the audio.
		/// </summary>
		public void Play()
		{
            AssertNonNullBuffer();

			_Buffer.Play();
		}

		/// <summary>
		/// Stops the audio playback.
		/// </summary>
		public void Stop()
		{
            AssertNonNullBuffer();

			_Buffer.Stop();
		}

		#endregion

		#region Finalization

		/// <summary>
		/// © 2005 IDesign Inc. Used by permission.
		/// </summary>
		protected bool Disposed
		{
			get
			{
				lock(this)
				{
					return m_Disposed;
				}
			}
		}

		/// <summary>
		/// © 2005 IDesign Inc. Used by permission.
		/// </summary>
		public void Dispose()
		{
			lock(this)
			{
				if(m_Disposed == false)
				{
					Cleanup();
					m_Disposed = true;
					GC.SuppressFinalize(this);
				}
			}
		}

		/// <summary>
		/// Cleans up resources.
		/// </summary>
		protected virtual void Cleanup()
		{
			if (_Buffer != null)
			{
				_Buffer.Dispose();
				_Buffer = null;
			}

			if (_Device != null)
			{
				_Device.Dispose();
				_Device = null;
			}
		}

		/// <summary>
		/// Dtor.
		/// </summary>
		~ManagedDirectSoundAudio()
		{
			try
			{
				Cleanup();
			}
			catch
			{
				//ignore any exceptions here
			}
		}
		#endregion

		private void _Buffer_Starting(object sender, EventArgs e)
		{
			try
			{
				if (Starting != null)
				{
					Starting(this, e);
				}
			}
			catch
			{}
		}

		private void _Buffer_Stopping(object sender, EventArgs e)
		{
			try
			{
				if (Stopping != null)
				{
					Stopping(this, e);
				}
			}
			catch
			{}
		}

		private void _Buffer_Pausing(object sender, EventArgs e)
		{
			try
			{
				if (Pausing != null)
				{
					Pausing(this, e);
				}
			}
			catch
			{}
		}
	}
}
