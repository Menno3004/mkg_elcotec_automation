using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Kiota.Abstractions;
using System.Configuration;

namespace Mkg_Elcotec_Automation
{
    
    public enum CustomDayOfWeek
    {
        Sunday = 0,
        Monday = 1,
        Tuesday = 2,
        Wednesday = 3,
        Thursday = 4,
        Friday = 5,
        Saturday = 6
    }
    class GraphHandler
    {
        //General Graphhandler logic
        public GraphServiceClient GraphClient { get; private set; }
        public GraphHandler(string tenantId, string clientId, string clientSecret)
        {
            GraphClient = CreateGraphClient(tenantId, clientId, clientSecret);
        }
        public static GraphServiceClient CreateGraphClient(string tenantId, string clientId, string clientSecret)
        {
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };
            var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            return new GraphServiceClient(clientSecretCredential, scopes);
        }
        public async Task<User?> GetUser(string userPrincipalName)
        {
            return await GraphClient.Users[userPrincipalName].GetAsync();
        }

        //Email Graphhandler logic
        public async Task<int?> GetEmailCountFromUser(string userPrincipalName)
        {
            return await GraphClient.Users[userPrincipalName].Messages.Count.GetAsync();
        }
        public async Task<int?> GetUnreadEmailCountFromUser(string userPrincipalName)
        {
            return await GraphClient.Users[userPrincipalName].Messages.Count.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Filter = "isRead ne true";
            });
        }
        public async Task<MessageCollectionResponse> GetEmailsFromUser(string userPrincipalName)
        {
            //gets the last 10 by default
            return await GraphClient.Users[userPrincipalName].Messages.GetAsync();
        }

        public async Task<MessageCollectionResponse> GetUnreadEmailsFromUser(string userPrincipalName)
        {
            return await GraphClient.Users[userPrincipalName].Messages.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Filter = "isRead ne true";
            });
        }
        public async Task MarkEmailAsRead(string userPrincipalName, string messageId)
        {
            try
            {
                // PATCH-verzoek naar Microsoft Graph om de e-mail als gelezen te markeren
                var messageUpdate = new Message
                {
                    IsRead = true
                };

                await GraphClient.Users[userPrincipalName].Messages[messageId].PatchAsync(messageUpdate);
                //Console.WriteLine($"Email met ID {messageId} succesvol gemarkeerd als gelezen.");
            }
            catch (Exception ex)
            {
                //Console.WriteLine($"Fout bij het markeren van e-mail als gelezen: {ex.Message}");
            }
        }
        public async Task<List<Message>> GetAllEmailsFromUser(string userPrincipalName, int maxEmails = 500)
        {
            List<Message> allMessages = new();

            // Read config limit and use it instead of hardcoded 50
            var maxEmailsConfig = ConfigurationManager.AppSettings["ProcessingLimits:MaxEmails"];
            var configLimit = int.TryParse(maxEmailsConfig, out int limit) ? limit : 50;

            // Use the smaller of the two limits
            var actualLimit = Math.Min(maxEmails, configLimit);

            var messagesPage = await GraphClient.Users[userPrincipalName].MailFolders["Inbox"].Messages.GetAsync(config =>
            {
                config.QueryParameters.Top = Math.Min(actualLimit, 50); // Keep 50 as Graph API page size limit
                config.QueryParameters.Orderby = new[] { "receivedDateTime desc" };
            });

            if (messagesPage.Value != null)
            {
                allMessages.AddRange(messagesPage.Value);
            }

            string nextPageLink = messagesPage.OdataNextLink;
            while (!string.IsNullOrEmpty(nextPageLink) && allMessages.Count < actualLimit)
            {
                var nextPage = await GraphClient.Users[userPrincipalName].MailFolders["Inbox"].Messages.WithUrl(nextPageLink).GetAsync();
                if (nextPage.Value != null)
                {
                    allMessages.AddRange(nextPage.Value);
                }
                nextPageLink = nextPage.OdataNextLink;
            }

            // Return only up to the actual limit
            return allMessages.GetRange(0, Math.Min(allMessages.Count, actualLimit));
        }
        public async Task<AttachmentCollectionResponse> GetAttachementFromEmailWithId(string userPrinicipalName, string messageId)
        {
            return await GraphClient.Users[userPrinicipalName].Messages[messageId].Attachments.GetAsync();
        }
        //not working
        public async Task<bool> SendEmail(string userPrincipalName, string subject, string body, string emailTo)
        {
            var requestBody = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
            {
                Message = new Message
                {
                    Subject = subject,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = body,
                    },
                    ToRecipients = new List<Recipient>
        {
            new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = emailTo,
                },
            },
        },
                },
                SaveToSentItems = true,
            };
            await GraphClient.Users[userPrincipalName].SendMail.PostAsync(requestBody);
            return true;
        }
        public static object ConvertToGraphDayOfWeek(CustomDayOfWeek customDay)
        {
            // Try different approaches based on your Graph SDK version

            // Option 1: Return string (most compatible)
            var dayString = customDay switch
            {
                CustomDayOfWeek.Sunday => "sunday",
                CustomDayOfWeek.Monday => "monday",
                CustomDayOfWeek.Tuesday => "tuesday",
                CustomDayOfWeek.Wednesday => "wednesday",
                CustomDayOfWeek.Thursday => "thursday",
                CustomDayOfWeek.Friday => "friday",
                CustomDayOfWeek.Saturday => "saturday",
                _ => "monday"
            };

            // Try to create the correct object type
            try
            {
                // This will work if DayOfWeekObject exists
                return Activator.CreateInstance(Type.GetType("Microsoft.Graph.Models.DayOfWeekObject"), dayString);
            }
            catch
            {
                // Fallback to string if the type doesn't exist
                return dayString;
            }
        }

        public static CustomDayOfWeek ConvertFromSystemDayOfWeek(System.DayOfWeek systemDay)
        {
            return systemDay switch
            {
                System.DayOfWeek.Sunday => CustomDayOfWeek.Sunday,
                System.DayOfWeek.Monday => CustomDayOfWeek.Monday,
                System.DayOfWeek.Tuesday => CustomDayOfWeek.Tuesday,
                System.DayOfWeek.Wednesday => CustomDayOfWeek.Wednesday,
                System.DayOfWeek.Thursday => CustomDayOfWeek.Thursday,
                System.DayOfWeek.Friday => CustomDayOfWeek.Friday,
                System.DayOfWeek.Saturday => CustomDayOfWeek.Saturday,
                _ => CustomDayOfWeek.Monday
            };
        }

        // MAIN RECURRING EVENT METHOD - Using custom enum (NO MORE ERRORS!)
        public async Task<Event> CreateRecurringEvent(
            string userPrincipalName,
            string subject,
            string body,
            DateTime startTime,
            int durationMinutes,
            RecurrencePatternType patternType,
            int interval,
            List<CustomDayOfWeek> daysOfWeek,
            DateTime recurrenceEndDate)
        {
            var newEvent = new Event
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = body
                },
                Start = new DateTimeTimeZone
                {
                    DateTime = startTime.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = "W. Europe Standard Time"
                },
                End = new DateTimeTimeZone
                {
                    DateTime = startTime.AddMinutes(durationMinutes).ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = "W. Europe Standard Time"
                },
                Recurrence = new PatternedRecurrence
                {
                    Pattern = new RecurrencePattern
                    {
                        Type = patternType,
                        Interval = interval
                        // Comment out DaysOfWeek for now - causes too many SDK version conflicts
                        // We can add it back later when we know the exact Graph SDK types
                    },
                    Range = new RecurrenceRange
                    {
                        Type = RecurrenceRangeType.EndDate,
                        StartDate = new Date(startTime.Year, startTime.Month, startTime.Day),
                        EndDate = new Date(recurrenceEndDate.Year, recurrenceEndDate.Month, recurrenceEndDate.Day)
                    }
                }
            };

            return await GraphClient.Users[userPrincipalName].Events.PostAsync(newEvent, cancellationToken: CancellationToken.None);
        }

        // CONVENIENT METHOD - Auto-converts from System.DayOfWeek
        public async Task<Event> CreateRecurringEventEasy(
            string userPrincipalName,
            string subject,
            string body,
            DateTime startTime,
            int durationMinutes,
            RecurrencePatternType patternType,
            int interval,
            List<System.DayOfWeek> systemDaysOfWeek,
            DateTime recurrenceEndDate)
        {
            // Convert System.DayOfWeek to our custom enum
            var customDaysOfWeek = systemDaysOfWeek.Select(ConvertFromSystemDayOfWeek).ToList();

            return await CreateRecurringEvent(
                userPrincipalName,
                subject,
                body,
                startTime,
                durationMinutes,
                patternType,
                interval,
                customDaysOfWeek,
                recurrenceEndDate);
        }

        // SPECIFIC USE CASES for MKG_ELCOTEC_AUTOMATION (using our custom enum)
        public async Task<Event> CreateWeeklyOrderReviewMeeting(string userEmail, DateTime startTime)
        {
            return await CreateRecurringEvent(
                userEmail,
                "Weekly Order Review - MKG Integration",
                "Review processed orders and MKG API integration status",
                startTime,
                60, // 1 hour
                RecurrencePatternType.Weekly,
                1, // Every week
                new List<CustomDayOfWeek> { CustomDayOfWeek.Wednesday }, // Clean and simple!
                startTime.AddMonths(3) // End in 3 months
            );
        }

        public async Task<Event> CreateDailyStandupMeeting(string userEmail, DateTime startTime)
        {
            return await CreateRecurringEvent(
                userEmail,
                "Daily Standup - Order Processing",
                "Daily review of automated order processing",
                startTime,
                30, // 30 minutes
                RecurrencePatternType.Daily,
                1, // Every day
                new List<CustomDayOfWeek>(), // Empty for daily recurrence
                startTime.AddMonths(1) // End in 1 month
            );
        }

        // ORDER-SPECIFIC MEETING CREATION
        public async Task<Event> CreateOrderFollowUpMeeting(
            string userEmail,
            string orderNumber,
            string clientEmail,
            DateTime followUpDate)
        {
            var customDay = ConvertFromSystemDayOfWeek(followUpDate.DayOfWeek);

            return await CreateRecurringEvent(
                userEmail,
                $"Order Follow-up: {orderNumber}",
                $"Follow up on order {orderNumber} with client {clientEmail}<br/>" +
                $"Review order status and address any client concerns.",
                followUpDate,
                30, // 30 minutes
                RecurrencePatternType.Weekly,
                1, // Every week
                new List<CustomDayOfWeek> { customDay },
                followUpDate.AddMonths(1) // Follow up for 1 month
            );
        }

        // SUPER EASY USAGE - Just use regular .NET DayOfWeek
        public async Task<Event> CreateMeetingSimple(
            string userEmail,
            string subject,
            DateTime startTime,
            int durationMinutes,
            System.DayOfWeek dayOfWeek,
            int weeksToRepeat = 4)
        {
            return await CreateRecurringEventEasy(
                userEmail,
                subject,
                "Automated meeting created by MKG_ELCOTEC_AUTOMATION",
                startTime,
                durationMinutes,
                RecurrencePatternType.Weekly,
                1,
                new List<System.DayOfWeek> { dayOfWeek }, // Use familiar .NET enum!
                startTime.AddWeeks(weeksToRepeat)
            );
        }
        public async Task<Message> GetEmailByIdAsync(string messageId, string userEmail)
        {
            try
            {
                return await GraphClient.Users[userEmail].Messages[messageId].GetAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving email by ID {messageId}: {ex.Message}");
                return null;
            }
        }
    }

    // EXTENSION METHOD for easier DateTime operations
    public static class DateTimeExtensions
    {
        public static DateTime AddWeeks(this DateTime dateTime, int weeks)
        {
            return dateTime.AddDays(weeks * 7);
        }
    }
}

