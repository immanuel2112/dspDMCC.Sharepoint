namespace dspDMCC.Sharepoint
{
    public class ADMTargetReportSegmentValue
    {
        public string TargetReportSegmentByFieldValue { get; set; }

        public string TargetReportSegmentLocation { get; set; }

        public override string ToString()
        {
            return "TargetReportSegmentByFieldValue: " + TargetReportSegmentByFieldValue +
                ", TargetReportSegmentLocation: " + TargetReportSegmentLocation;
        }
    }
}
