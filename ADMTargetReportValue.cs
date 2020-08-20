using System.Collections.Generic;

namespace dspDMCC.Sharepoint
{
    public class ADMTargetReportValue
    {
        public string TargetReportID { get; set; }

        public string TargetReport { get; set; }

        public string TargetReportType { get; set; }

        public string TargetReportLocation { get; set; }

        public string TargetReportSegmentByField { get; set; }

        public List<ADMTargetReportSegmentValue> TargetReportSegmentValues { get; set; }

        public override string ToString()
        {
            return "TargetReportID: " + TargetReportID +
                ", TargetReport: " + TargetReport +
                ", TargetReportType: " + TargetReportType +
                ", TargetReportLocation: " + TargetReportLocation +
                ", TargetReportSegmentByField: " + TargetReportSegmentByField +
                ", TargetReportSegmentValues: " + TargetReportSegmentValues;
        }
    }
}
