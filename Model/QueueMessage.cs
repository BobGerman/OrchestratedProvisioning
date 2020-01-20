using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace OrchestratedProvisioning.Model
{
    class QueueMessage
    {
        // If a property is missing, StringEnumConverter will set the 1st value
        // Therefore be sure the 1st enumeration is unknown so we can detect this
        public enum Command
        {
            unknown, provisionModernTeamSite, applyProvisioningTemplate, createTeam
        }

        public enum ResultCode
        {
            unknown, success, warning, failure
        }

        [JsonConverter(typeof(StringEnumConverter))]
        public Command command { get; set; }
        public string template { get; set; }
        public string alias { get; set; }
        public string requestId { get; set; }
        public string displayName { get; set; }
        public string description { get; set; }
        public string owner { get; set; }
        public bool isPublic { get; set; }
        public string groupId { get; set; }
        [JsonConverter(typeof(StringEnumConverter))]
        public ResultCode resultCode { get; set; }
        public string resultMessage { get; set; }

        public string Serialize()
        {
            return JsonConvert.SerializeObject(this);
        }

        public static QueueMessage NewFromJson(string val)
        {
            return JsonConvert.DeserializeObject<QueueMessage>(val);
        }
    }
}
