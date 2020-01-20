namespace OrchestratedProvisioning.Model
{
    // Various constants used in the application
    static public class Constants
    { 
        // Queue names
        public const string RequestQueueName = "provisioning-request-queue";
        public const string CompletionQueueName = "provisioning-completion-queue";

        // Graph async calls
        public const int RetryInterval = 5000;      // 5 seconds
        public const int RetryMax = 120;            // ~10 minutes of retries @ 5sec interval
    }
}
