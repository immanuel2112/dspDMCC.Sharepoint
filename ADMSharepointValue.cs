using CranBerry.Framework.Data;
using System.Collections.Generic;

namespace dspDMCC.Sharepoint
{
    public class ADMSharepointValue
    {
        public IDataService Sys { get; set; }

        public string LoadManagerId { get; set; }

        public string Wave { get; set; }

        public string Processarea { get; set; }

        public string Object { get; set; }

        public string Target { get; set; }

        public string WaveProcessareaObjectTargetID { get; set; }

        public string LoadCycle { get; set; }

        public string InitialDelta { get; set; }

        public string Version { get; set; }

        public SharepointValue SharepointValue { get; set; }

        public List<ADMTargetReportValue> TargetReportValues { get; set; }

        public string ADMReportPath { get; set; }
    }
}
