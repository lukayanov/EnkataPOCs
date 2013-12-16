using Microsoft.Exchange.WebServices.Data;
using NLog;
using System;
using System.Configuration;
using System.Net;

namespace ExchangeTracker
{
    class ExchangeWatcher : IDisposable
    {
        readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private StreamingSubscriptionConnection _connection;
        private ExchangeService _service;
        private StreamingSubscription _streamingsubscription;

        public ExchangeWatcher()
        {
            _logger.Debug("Starting");
            ConnectToServer();
            ListAllFolders();
//            ListInboxFolder();
            _logger.Info("Started");
        }

        private void ListInboxFolder()
        {
            var inbox = Folder.Bind(_service, WellKnownFolderName.Inbox);
            var itemView = new ItemView(int.MaxValue);
            var searchResults = _service.FindItems(inbox.Id, itemView);
            foreach (var item in searchResults)
            {
                var message = item as EmailMessage;
                if (message == null)
                    throw new Exception();
                _logger.Info("New mail from: {0}, Subject: {1}", message.From.Name, message.Subject);
            }
        }

        public void Dispose()
        {
            _logger.Debug("Disposing");
            UnsubscribeEvents();
            _logger.Info("Disposed");
        }

        void SubscribeToEvents()
        {
            _streamingsubscription = _service.SubscribeToStreamingNotifications(
                new FolderId[] { WellKnownFolderName.Inbox, WellKnownFolderName.Calendar},
                EventType.NewMail,
                EventType.Created,
                EventType.Deleted,
                EventType.FreeBusyChanged,
                EventType.Deleted
                );

            _connection = new StreamingSubscriptionConnection(_service, 1);

            _connection.AddSubscription(_streamingsubscription);
            _connection.OnNotificationEvent += OnEvent;
            _connection.OnSubscriptionError += OnError;
            _connection.OnDisconnect += OnDisconnect;
            _connection.Open();

            _logger.Info("Subscription completed");
        }

        void UnsubscribeEvents()
        {

            _connection.RemoveSubscription(_streamingsubscription);
            _connection.OnNotificationEvent -= OnEvent;
            _connection.OnSubscriptionError -= OnError;
            _connection.OnDisconnect -= OnDisconnect;
            _connection.Close();
            _logger.Info("Unsubscription completed");
        }

        static string EwsEndpoint
        {
            get
            {
                return ConfigurationManager.AppSettings["EWSEndpoint"];
            }
        }

        private void ConnectToServer()
        {
            _service = new ExchangeService(ExchangeVersion.Exchange2010_SP1)
            {
//                Credentials = new WebCredentials(CredentialCache.DefaultCredentials),
                Credentials = new WebCredentials("testadmin", "Enkata!!"),
                Url = new Uri(EwsEndpoint)
            };
            _connection = new StreamingSubscriptionConnection(_service, 1);
            SubscribeToEvents();
        }

        private void OnDisconnect(object sender, SubscriptionErrorEventArgs args)
        {
            try
            {
                var connection = (StreamingSubscriptionConnection)sender;
                _logger.Info("Connection restored");
                connection.Open();
            }
            catch (Exception ex)
            {
                _logger.ErrorException("Failed to reconnect", ex);
            }
        }

        private void OnError(object sender, SubscriptionErrorEventArgs args)
        {
            _logger.ErrorException("Error", args.Exception);
        }

        private void OnEvent(object sender, NotificationEventArgs args)
        {
            foreach (var notification in args.Events)
            {
                var itemEvent = notification as ItemEvent;
                if (itemEvent != null)
                {
                    OnItemEvent(itemEvent);
                }
            }
        }

        private void OnItemEvent(ItemEvent notification)
        {
            _logger.Info("Event type: {0}, timestamp:{1}, Id: {2}", notification.EventType.ToString(), notification.TimeStamp, notification.ItemId);
            switch (notification.EventType)
            {
                case EventType.NewMail:
                    OnNewMail(notification);
                    break;
                case EventType.Created:
                    _logger.Info("Item or folder created");
                    break;
                case EventType.Deleted:
                    _logger.Info("Item or folder deleted");
                    break;
            }
        }

        private void OnNewMail(ItemEvent notification)
        {
            _logger.Info("New mail received");
            var mail = ((EmailMessage)Item.Bind(_service, notification.ItemId));
            _logger.Info("New mail from: {0}, Subject: {1}", mail.From.Name, mail.Subject);
            _logger.Info("Mail content: {0}", mail.Body);
        }

        private void ListAllFolders()
        {
            var view = new FolderView(int.MaxValue)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly) {FolderSchema.DisplayName}
            };
            SearchFilter searchFilter = new SearchFilter.IsGreaterThan(FolderSchema.TotalCount, 0);
            view.Traversal = FolderTraversal.Deep;

            // Send the request to search the mailbox and get the results.
            var findFolderResults = _service.FindFolders(WellKnownFolderName.MsgFolderRoot, searchFilter, view);

            // Process each item.
            foreach (var myFolder in findFolderResults.Folders)
            {
                if (myFolder is SearchFolder)
                {
                    _logger.Info("Search folder: {0}", (myFolder as SearchFolder).DisplayName);
                }

                else if (myFolder is ContactsFolder)
                {
                    _logger.Info("Contacts folder: {0}", (myFolder as ContactsFolder).DisplayName);
                }

                else if (myFolder is TasksFolder)
                {
                    _logger.Info("Tasks folder: {0}", (myFolder as TasksFolder).DisplayName);
                }

                else if (myFolder is CalendarFolder)
                {
                    _logger.Info("Calendar folder: {0}", (myFolder as CalendarFolder).DisplayName);
                }
                else
                {
                    // Handle a generic folder.
                    _logger.Info("Folder: {0}", myFolder.DisplayName);
                }
            }

            // Determine whether there are more folders to return.
            if (findFolderResults.MoreAvailable)
            {

            }            
        }
    }
}
