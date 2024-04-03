using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IndiaEvents.Models.Models.Webhook
{
    public class WebhookModels
    {
    }
    public class Webhook
    {
        public string? smartsheetHookResponse { get; set; }
        public string challenge { get; set; }
        public string webhookId { get; set; }

    }


    public class Event
    {
        public string objectType { get; set; }
        public string eventType { get; set; }
        public long rowId { get; set; }
        public long columnId { get; set; }
        public long userId { get; set; }
        public DateTime timestamp { get; set; }
    }

    public class Root
    {
        public long webhookId { get; set; }
        public string challenge { get; set; }
        public string nonce { get; set; }
        public DateTime timestamp { get; set; }
        public string scope { get; set; }
        public long scopeObjectId { get; set; }
        public List<Event> events { get; set; }
    }

    //public class SmartsheetWebhookPayload
    //{
    //    public string? Event { get; set; }
    //    public string? Changes { get; set; }
    //} 
    //public class SmartsheetWebhookPayload
    //{
    //    public SmartsheetEvent Event { get; set; }
    //    public List<SmartsheetChange> Changes { get; set; }
    //}

    //public class SmartsheetEvent
    //{
    //    public long ObjectId { get; set; }
    //}

    //public class SmartsheetChange
    //{
    //    public SmartsheetCell Cell { get; set; }
    //}

    //public class SmartsheetCell
    //{
    //    public string Value { get; set; }
    //}

}
