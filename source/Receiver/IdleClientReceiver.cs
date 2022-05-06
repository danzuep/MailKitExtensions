using MailKit;
using MailKit.Security;
using MailKit.Net.Imap;
using System.Net;
using System.Net.Sockets;
using Microsoft.Extensions.Logging;
using Zue.Common;

#region Author Credit Attribution Statement
//
// IdleClient.cs
//
// Author: Jeffrey Stedfast <jeff@xamarin.com>
//
// Copyright (c) 2014-2020 Jeffrey Stedfast
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
// https://github.com/jstedfast/MailKit/blob/master/Documentation/Examples/ImapIdleExample.cs
#endregion

namespace MailKitExtensions.Receiver
{
    public delegate ValueTask MessagesArrived(IList<IMessageSummary> messages);

    public sealed class IdleClientReceiver : EmailReceiver
	{
		public const string DateTimeFormatLong = "yyyy'-'MM'-'dd' 'HH':'mm':'ss'.'fff' 'K"; //:O has fffffffK
		public const string DateTimeFormat = "yyyy'-'MM'-'dd' 'HH':'mm':'ss"; //:s has 'T'
		public const string DateOnlyFormat = "yyyy'-'MM'-'dd";
		public const string TimeOnlyFormat = "HH':'mm':'ss";
		public const string TimeOnlyFormatLong = "HH':'mm':'ss'.'fff";

		#region Private Fields
		const int _maxRetries = 3;
		private bool _messagesArrived;
		private bool _processingMessages;
		private CancellationTokenSource _cancel = new CancellationTokenSource();
        private CancellationTokenSource _done = new CancellationTokenSource();
		private IList<IMessageSummary> _messageCache = new List<IMessageSummary>();
		MessagesArrived ProcessMessageSummariesAsync;
		#endregion

		#region Initialisation
		public IdleClientReceiver(NetworkCredential credential, MessagesArrived messagesArrivedMethod,
			string imapHost, string folderName = "INBOX", bool connect = true, bool useLogger = false)
			: base(credential, imapHost, folderName, connect, useLogger)
			=> ProcessMessageSummariesAsync = messagesArrivedMethod;
		#endregion

		public async Task RunAsync(CancellationToken ct = default)
		{
			try
			{
				var mailFolder = ConnectFolder(true, ct);

				mailFolder.CountChanged += OnCountChanged;
				mailFolder.MessageExpunged += OnMessageExpunged;
				mailFolder.MessageFlagsChanged += OnMessageFlagsChanged;

				_messagesArrived = mailFolder.Count > 0;
				_logger.LogInformation("Email idle client started watching '{0}' ({1}) {2}.",
					MailFolder?.FullName, MailFolder?.Count, DateTime.Now.ToString(DateTimeFormat));
				await IdleAsync(ct);
				_logger.LogInformation("Email idle client finished watching '{0}' ({1}) {2}.",
					MailFolder?.FullName, MailFolder?.Count, DateTime.Now.ToString(DateTimeFormat));
				_messageCache.Clear();

				mailFolder.MessageFlagsChanged -= OnMessageFlagsChanged;
				mailFolder.MessageExpunged -= OnMessageExpunged;
				mailFolder.CountChanged -= OnCountChanged;

				var mailFolderName = _folderName;
				base.Disconnect();

                if (_folderName != mailFolderName && !string.IsNullOrWhiteSpace(mailFolderName))
                {
                    _logger.LogError("MailFolder name had changed from '{0}' to '{1}'!", mailFolderName, _folderName);
                    _ = OpenFolder(mailFolderName, false);
                }
            }
			catch (ServiceNotConnectedException ex)
			{
				_logger.LogWarning(ex, ex.Message);
				return;
			}
			catch (OperationCanceledException)
			{
				_logger.LogDebug("Initial fetch or idle task in IdleClient was cancelled.");
				return;
			}
			catch (InvalidOperationException ex)
			{
				_logger.LogError(ex, "IMAP client is busy.");
				return;
			}
			catch (Exception ex)
			{
				_logger.LogError(ex, "IMAP idle client failed while fetching initial messages.");
				throw;
			}
		}

		async Task<int> ProcessMessagesAsync()
		{
			IList<IMessageSummary> fetched = new List<IMessageSummary>();
			int retryCount = 0;

			do
			{
				try
				{
					if (!_cancel.IsCancellationRequested)
					{
						ConnectFolder(true);
						int startIndex = _messageCache.Count;
						_logger.LogTrace("'{0}' ({1}): Fetching new message arrivals, starting from {2}.",
							MailFolder.FullName, MailFolder.Count, startIndex);
						if (startIndex > MailFolder.Count)
							startIndex = MailFolder.Count;
						lock (MailFolder.SyncRoot)
						{
							fetched = MailFolder.Fetch(startIndex, -1,
								MessageSummaryItems.UniqueId, _cancel.Token);
						}
						AddFetchedMessages(fetched);
						_logger.LogDebug("Downloading {0} from '{1}' ({2}) at {3:HH':'mm':'ss'.'fff}. ID{4}: {5}",
							fetched.Count, MailFolder.FullName, MailFolder.Count, DateTime.Now,
							fetched.Count.S(), fetched.Select(m => m.UniqueId).ToEnumeratedString());
						await ProcessMessageSummariesAsync(fetched);
					}
					else
						_logger.LogInformation("Service was cancelled during fetch.");
					break;
				}
				catch (ImapCommandException ex)
				{
					// command exceptions often result in the client getting disconnected
					_logger.LogWarning(ex, "Client request to examine 'INBOX' was denied, server unavailable. Reconnecting.");
					Reconnect();
				}
				catch (ImapProtocolException)
				{
					// protocol exceptions often result in the client getting disconnected
					_logger.LogDebug("Imap protocol exception, reconnecting.");
					Reconnect();
				}
				catch (IOException)
				{
					// I/O exceptions always result in the client getting disconnected
					_logger.LogDebug("Imap I/O exception, reconnecting.");
					Reconnect();
				}
				catch (OperationCanceledException) // includes TaskCanceledException
				{
					_logger.LogDebug("Fetch task cancelled in email IdleClient.");
					RemoveFetchedMessages(fetched);
				}
				catch (Exception ex)
				{
					_logger.LogWarning(ex, "Imap idle client failed to fetch new messages.");
					RemoveFetchedMessages(fetched);
					throw;
				}
				finally
				{
					retryCount++;
					if (retryCount >= _maxRetries)
						RemoveFetchedMessages(fetched);
				}
			} while (retryCount < _maxRetries);

			int processedCount = fetched.Count;
			if (fetched.Count > 0)
				_logger.LogTrace("{0} message{s} processed.", processedCount, processedCount.S());

			return processedCount;
		}

		private void AddFetchedMessages(IList<IMessageSummary> items, bool allowReprocess = true)
		{
			if (items != null)
				foreach (var item in items)
					if (allowReprocess || !_messageCache.Contains(item))
						_messageCache.Add(item);
		}

		private void RemoveFetchedMessages(IList<IMessageSummary> items)
        {
			if (items != null)
				foreach (var item in items)
					_messageCache.Remove(item);
		}

		async Task WaitForNewMessagesAsync()
		{
			do
			{
				try
				{
					if (_imapClient.Capabilities.HasFlag(ImapCapabilities.Idle))
					{
						// Note: IMAP servers are only supposed to drop the connection after 30 minutes, so normally
						// we'd IDLE for a max of, say, ~29 minutes... but GMail seems to drop idle connections after
						// about 10 minutes, so we'll only idle for 9 minutes.
						_done = new CancellationTokenSource(TimeSpan.FromMinutes(9));
						var cancel = CancellationTokenSource.CreateLinkedTokenSource(_cancel.Token);

						_logger.LogTrace("Idle wait task started {0:HH':'mm':'ss'.'fff}.", DateTime.Now);
						ConnectFolder(true, cancel.Token);
						await _imapClient.IdleAsync(_done.Token, cancel.Token);
						//_logger.LogTrace("Idle wait task finished, done: {0}, cancel: {1}, {2:HH':'mm':'ss'.'fff}.",
						//	_done.Token.IsCancellationRequested, cancel.Token.IsCancellationRequested, DateTime.Now);

						if (_done.IsCancellationRequested && _messageCache.Count > 0 && MailFolder?.Count > 0 && MailFolder.Count != _messageCache.Count)
						{
							_logger.LogInformation("'{0}' ({1}) count resynchronised from {2} back to 0.",
								MailFolder?.FullName, MailFolder?.Count, _messageCache.Count);
							_messageCache.Clear();
						}
						else if (_done.IsCancellationRequested)
						{
							_logger.LogTrace("'{0}' ({1}) idle client count is {2} on reset.",
								MailFolder?.FullName, MailFolder?.Count, _messageCache.Count);
						}

						_done.Dispose();
						_done = null;
					}
					else
					{
						// Use SMTP NOOP commands to simulate the IMAP idle capability, but don't spam it.
						await Task.Delay(TimeSpan.FromMinutes(1), _cancel.Token);
						await _imapClient.NoOpAsync(_cancel.Token);
					}
				}
				catch (ServiceNotConnectedException)
				{
					_logger.LogInformation("IMAP service not connected, reconnecting.");
					Reconnect();
				}
				catch (ImapProtocolException)
				{
					// protocol exceptions often result in the client getting disconnected
					_logger.LogInformation("IMAP protocol exception, reconnecting.");
					Reconnect();
				}
				catch (IOException)
				{
					// I/O exceptions always result in the client getting disconnected
					_logger.LogInformation("IMAP I/O exception, reconnecting.");
					Reconnect();
				}
				catch (ImapCommandException ex)
				{
					// command exceptions often result in the client getting disconnected
					_logger.LogWarning(ex, "Client request to examine 'INBOX' was denied, server unavailable. Reconnecting.");
					Reconnect();
				}
				catch (ObjectDisposedException)
				{
					_logger.LogWarning("Cancellation token object disposed.");
					_cancel.Token.ThrowIfCancellationRequested();
				}
				catch (NullReferenceException ex)
				{
					_logger.LogWarning(ex, "Cancellation token object disposed, null reference.");
					_cancel.Token.ThrowIfCancellationRequested();
				}
				catch (InvalidOperationException ex)
				{
					_logger.LogError(ex, "IMAP client is being accessed by multiple threads.");
					_cancel.Cancel(false);
				}
				catch (OperationCanceledException) // includes TaskCanceledException
				{
					_logger.LogTrace("Idle wait task cancelled.");
					_done?.Cancel(false);
				}
				catch (Exception ex)
				{
					_logger.LogError(ex, "IMAP idle client failed while waiting for new messages.");
					throw;
				}
			} while (_done != null &&
				!_done.IsCancellationRequested &&
				!_cancel.IsCancellationRequested);
		}

		async Task IdleAsync(CancellationToken ct = default)
		{
			if (ct != default)
				_cancel = CancellationTokenSource.CreateLinkedTokenSource(ct);
			do
			{
				try
				{
					if (!_messagesArrived)
						await WaitForNewMessagesAsync();
					if (_messagesArrived && !_processingMessages)
					{
						_processingMessages = true;
						_messagesArrived = false;
						await ProcessMessagesAsync();
						_processingMessages = false;
					}
				}
				catch (SocketException ex) // thrown from Reconnect() after IOException
				{
					_logger.LogWarning(ex, "Error re-thrown from Reconnect() after IOException. {0}", ex.Message);
				}
				catch (OperationCanceledException) // includes TaskCanceledException
				{
					_logger.LogDebug("Idle task cancelled.");
					break;
				}
				catch (AuthenticationException ex)
				{
					_logger.LogWarning(ex, "{0}: Stopping idle client service.", ex.GetType().Name);
					break;
				}
				catch (Exception ex)
				{
					_logger.LogWarning(ex, "Imap idle client failed while idleing or processing messages.");
					throw;
				}
			}
			while (!_cancel.IsCancellationRequested);
		}

		// Note: the CountChanged event will fire when new messages arrive in the folder and/or when messages are expunged.
		// Keep track of changes to the number of messages in the folder (this is how we'll tell if new messages have arrived).
		void OnCountChanged(object sender, EventArgs e)
		{
			var folder = sender as ImapFolder;

			// Note: because we are keeping track of the MessageExpunged event and updating our
			// 'messages' list, we know that if we get a CountChanged event and folder.Count is
			// larger than messages.Count, then it means that new messages have arrived.
			using (_logger.BeginScope("OnCountChanged"))
			{
				int folderCount = folder?.Count ?? 0;
				int changeCount = folderCount - _messageCache.Count;
				if (changeCount > 0)
				{
					_logger.LogTrace("[folder] '{0}' message count increased by {1} ({2} to {3}) at {4:HH':'mm':'ss'.'fff}.",
						folder.FullName, changeCount, _messageCache.Count, folderCount, DateTime.Now);

					// The ImapFolder is not re-entrant, so fetch the summaries later
					MailFolder = folder;
					_messagesArrived = true;
					_done?.Cancel();
				}
			}
		}

		// Keep track of messages being expunged so that when the CountChanged event fires, we can tell if it's
		// because new messages have arrived vs messages being removed (or some combination of the two).
		void OnMessageExpunged(object sender, MessageEventArgs e)
		{
			if (e.Index < _messageCache.Count)
			{
				// remove the locally cached message at e.Index.
				_messageCache.RemoveAt(e.Index);

				if (sender is ImapFolder folder)
					_logger.LogTrace("{0}: message index {1} expunged at {2:HH':'mm':'ss'.'fff}.",
						folder.FullName, e.Index, DateTime.Now);
			}
		}

		// keep track of flag changes
		void OnMessageFlagsChanged(object sender, MessageFlagsChangedEventArgs e)
		{
			if (sender is ImapFolder folder)
				_logger.LogTrace("{0}: flags have changed for message index {1} ({2}).",
					folder.FullName, e.Index, e.Flags);
		}

		public override void Dispose()
		{
			base.Dispose();
            _cancel?.Cancel(false);
			//_cancel?.Dispose(); //Careful, the cancel token is linked to the service token
		}
	}
}
