using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel.Channels;
using System.Text;
using System.Threading.Tasks;
using Xrm.Oss.FluentQuery;
using Xrm.Oss.XTL.Interpreter;

namespace Xrm.Oss.PowerKanban
{

    [DataContract]
    public class CreateNotificationConfig
    {
        [DataMember(Name = "parentLookupName")]
        public string ParentLookupName { get; set; }

        [DataMember(Name = "subscriptionLookupName")]
        public string SubscriptionLookupName { get; set; }

        [DataMember(Name = "notificationLookupName")]
        public string NotificationLookupName { get; set; }

        [DataMember(Name = "notifyCurrentUser")]
        public bool NotifyCurrentUser { get; set; }

        [DataMember(Name = "capturedFields")]
        public List<string> CapturedFields { get; set; }

        [DataMember(Name = "xtlCondition")]
        public string XtlCondition { get; set; }

        [DataMember(Name = "globalNotificationConfig")]
        public Dictionary<string, string> GlobalNotificationConfig { get; set; }

        [DataMember(Name = "messageConfig")]
        public Dictionary<string, string> MessageConfig { get; set; }

        [DataMember(Name = "emailLocaleConfig")]
        public Dictionary<string, string> EmailLocaleConfig { get; set; }
    }
    
    [DataContract]
    public class GlobalNotificationActionPayload
    {
        [DataMember(Name = "entityName")]
        public string EntityName { get; set; }

        [DataMember(Name = "message")]
        public string Message { get; set; }
    }

    public enum EventType {
        Update = 863910000,
        Create = 863910001,
        Assign = 863910002,
        Delete = 863910003,
        UserMention = 863910004
    }

    public class CreateNotification : IPlugin
    {
        private readonly Lazy<CreateNotificationConfig> config;

        public CreateNotification(string unsecure, string secure)
        {
            config = new Lazy<CreateNotificationConfig>(() => JsonDeserializer.Parse<CreateNotificationConfig>(unsecure));
        }

        public Entity GetTarget(IPluginExecutionContext context)
        {
            if (context.InputParameters.ContainsKey("Target"))
            {
                return context.InputParameters["Target"] as Entity;
            }

            return null;
        }

        public EntityReference GetTargetRef(IPluginExecutionContext context)
        {
            if (context.InputParameters.ContainsKey("Target"))
            {
                return context.InputParameters["Target"] as EntityReference;
            }

            return null;
        }

        public EventType GetEventType(IPluginExecutionContext context)
        {
            switch (context.MessageName.ToLowerInvariant())
            {
                case "create":
                    return EventType.Create;
                case "update":
                    return EventType.Update;
                case "assign":
                    return EventType.Assign;
                case "delete":
                    return EventType.Delete;
                default:
                    return EventType.UserMention;
            }
        }

        private T GetValue<T> (string attribute, Entity target, EntityImageCollection preImages)
        {
            if (string.IsNullOrEmpty(attribute))
            {
                return default(T);
            }

            return target.Contains(attribute)
                ? target.GetAttributeValue<T>(attribute)
                : preImages.Select(p => p.Value.GetAttributeValue<T>(config.Value.ParentLookupName)).FirstOrDefault();
        }

        private bool CheckExecutionCondition(Entity target, IOrganizationService service, ITracingService tracing)
        {
            var condition = config.Value.XtlCondition;

            if (string.IsNullOrEmpty(condition))
            {
                return true;
            }

            var result = new XTLInterpreter(condition, target, new OrganizationConfig(), service, tracing).Produce();

            if (!bool.TryParse(result, out var resultBool)) {
                tracing.Trace("When using the xtlCondition, make sure its result is a boolean");
                return false;
            }

            return resultBool;
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            var context = serviceProvider.GetService(typeof(IPluginExecutionContext)) as IPluginExecutionContext;
            var crmTracing = serviceProvider.GetService(typeof(ITracingService)) as ITracingService;
            var serviceFactory = serviceProvider.GetService(typeof(IOrganizationServiceFactory)) as IOrganizationServiceFactory;
            var service = serviceFactory.CreateOrganizationService(context.UserId);

            var target = GetTarget(context);
            var targetRef = GetTargetRef(context);

            if (target == null && targetRef == null)
            {
                return;
            }

            var shouldExecute = CheckExecutionCondition(target, service, crmTracing);

            if (!shouldExecute)
            {
                crmTracing.Trace("Execution condition not met, aborting");
                return;
            }

            var attributes = target != null
                ? target.Attributes.Keys.ToList()
                : null;

            var filteredAttributes = config.Value.CapturedFields != null
                ? attributes.Where(a => config.Value.CapturedFields.Any(f => string.Equals(a, f, StringComparison.InvariantCultureIgnoreCase))).ToList()
                : attributes;

            var eventData = new EventData
            {
                UpdatedFields = filteredAttributes,
                EventRecordReference = target?.ToEntityReference() ?? targetRef
            };

            var eventTarget = string.IsNullOrEmpty(config.Value.ParentLookupName) && eventData.EventRecordReference.Id != Guid.Empty
                ? eventData.EventRecordReference
                : GetValue<EntityReference>(config.Value.ParentLookupName, target, context.PreEntityImages);

            if (eventTarget == null) {
                crmTracing.Trace("Failed to find parent, exiting");
                return;
            }

            var subscriptionsQuery = service.Query("oss_subscription")
                .Where(e => e
                    .Attribute(a => a
                        .Named(config.Value.SubscriptionLookupName)
                        .Is(ConditionOperator.Equal)
                        .To(eventTarget.Id)
                    )
                    .Attribute(a => a
                        .Named("statecode")
                        .Is(ConditionOperator.Equal)
                        .To(0)
                    )
                )
                .IncludeColumns("ownerid", "oss_emailnotificationsenabled", "oss_emailnotificationssender")
                .Link(l => l
                    .FromEntity("oss_subscription")
                    .ToEntity("systemuser")
                    .FromAttribute("ownerid")
                    .ToAttribute("systemuserid")
                    .With.LinkType(JoinOperator.LeftOuter)
                    .Link(l2 => l2
                        .FromEntity("systemuser")
                        .ToEntity("usersettings")
                        .FromAttribute("systemuserid")
                        .ToAttribute("systemuserid")
                        .With.LinkType(JoinOperator.LeftOuter)
                        .With.Alias("usersettings")
                        .IncludeColumns("localeid")
                    )
                );

            if (!config.Value.NotifyCurrentUser) {
                subscriptionsQuery.AddCondition(
                    (a => a
                        .Named("ownerid")
                        .Is(ConditionOperator.NotEqual)
                        .To(context.UserId)
                    )
                );
            }
                
            var subscriptions = subscriptionsQuery.RetrieveAll();

            var serializedNotification = JsonSerializer.Serialize(eventData);
            var eventType = GetEventType(context);

            var messageConfig = config.Value.MessageConfig ?? new Dictionary<string, string>();

            var messages = subscriptions.Select(s => s.GetAttributeValue<AliasedValue>("usersettings.localeid")?.Value as int?)
                .Select(locale => locale != null ? locale.Value.ToString() : "default")
                .Distinct()
                .ToDictionary(
                    (k) => k,
                    (k) => messageConfig.ContainsKey(k)
                        ? TokenMatcher.ProcessTokens(messageConfig[k], target, new OrganizationConfig(), service, crmTracing)
                        : (messageConfig.ContainsKey("default") ? TokenMatcher.ProcessTokens(messageConfig["default"], target, new OrganizationConfig(), service, crmTracing) : null)
                );

            subscriptions.ForEach(subscription =>
            {
                var localeCode = subscription.GetAttributeValue<AliasedValue>("usersettings.localeid")?.Value as int?;
                var locale = localeCode != null ? localeCode.Value.ToString() : "default";

                var message = messages.ContainsKey(locale) ? messages[locale] : null;
                var user = subscription.GetAttributeValue<EntityReference>("ownerid");

                var notification = new Entity
                {
                    LogicalName = "oss_notification",
                    Attributes = {
                        ["oss_subscriptionid"] = subscription.ToEntityReference(),
                        ["ownerid"] = user,
                        ["oss_event"] = new OptionSetValue((int) eventType),
                        [config.Value.NotificationLookupName] = eventTarget,
                        ["oss_data"] = serializedNotification,
                        ["oss_text"] = message
                    }
                };

                if (config.Value.GlobalNotificationConfig != null)
                {
                    var messageKey = config.Value.GlobalNotificationConfig.Keys.FirstOrDefault(k => k == locale) ?? config.Value.GlobalNotificationConfig.Keys.FirstOrDefault(k => k == "default");
                    
                    if (messageKey != null)
                    {
                        var notificationMessage = config.Value.GlobalNotificationConfig[messageKey];

                        var payload = new GlobalNotificationActionPayload
                        {
                            EntityName = eventTarget?.LogicalName,
                            Message = notificationMessage
                        };

                        notification["oss_globalnotificationactionpayload"] = JsonSerializer.Serialize(payload);
                    }
                }

                notification.Id = service.Create(notification);

                var emailNotificationsEnabled = subscription.GetAttributeValue<bool>("oss_emailnotificationsenabled");

                if (!emailNotificationsEnabled)
                {
                    return;
                }

                SendEmailNotification(message, locale, user, subscription, notification, service);
            });
        }

        private void SendEmailNotification(string message, string locale, EntityReference user, Entity subscription, Entity notification, IOrganizationService service)
        {
            var emailNotificationsSenderText = subscription.GetAttributeValue<string>("oss_emailnotificationssender");
            var sender = string.IsNullOrEmpty(emailNotificationsSenderText) ? user : JsonDeserializer.Parse<EntityReference>(emailNotificationsSenderText);

            var email = new Entity
            {
                LogicalName = "email",
                ["from"] = new EntityCollection(new List<Entity>
                    {
                        new Entity {
                            LogicalName = "activityparty",
                            ["partyid"] = sender
                        }
                    }),
                ["to"] = new EntityCollection(new List<Entity>
                    {
                        new Entity {
                            LogicalName = "activityparty",
                            ["partyid"] = user
                        }
                    }),
                ["regardingobjectid"] = notification.ToEntityReference(),
                ["subject"] = "New D365 Notification",
                ["description"] = message
            };

            if (config.Value.EmailLocaleConfig != null && config.Value.EmailLocaleConfig.ContainsKey(locale))
            {
                email["oss_htmltemplateuniquename"] = config.Value.EmailLocaleConfig[locale];
            }
            else if (config.Value.EmailLocaleConfig != null && config.Value.EmailLocaleConfig.ContainsKey("default"))
            {
                email["oss_htmltemplateuniquename"] = config.Value.EmailLocaleConfig["default"];
            }

            email.Id = service.Create(email);

            var sendRequest = new SendEmailRequest
            {
                EmailId = email.Id,
                IssueSend = true
            };

            service.Execute(sendRequest);
        }
    }
}
