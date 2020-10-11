using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class DIGSIG
    {
        public static readonly XNamespace digsig =
            "http://schemas.microsoft.com/office/2006/digsig";

        public static readonly XName ApplicationVersion = digsig + "ApplicationVersion";
        public static readonly XName ColorDepth = digsig + "ColorDepth";
        public static readonly XName HorizontalResolution = digsig + "HorizontalResolution";
        public static readonly XName ManifestHashAlgorithm = digsig + "ManifestHashAlgorithm";
        public static readonly XName Monitors = digsig + "Monitors";
        public static readonly XName OfficeVersion = digsig + "OfficeVersion";
        public static readonly XName SetupID = digsig + "SetupID";
        public static readonly XName SignatureComments = digsig + "SignatureComments";
        public static readonly XName SignatureImage = digsig + "SignatureImage";
        public static readonly XName SignatureInfoV1 = digsig + "SignatureInfoV1";
        public static readonly XName SignatureProviderDetails = digsig + "SignatureProviderDetails";
        public static readonly XName SignatureProviderId = digsig + "SignatureProviderId";
        public static readonly XName SignatureProviderUrl = digsig + "SignatureProviderUrl";
        public static readonly XName SignatureText = digsig + "SignatureText";
        public static readonly XName SignatureType = digsig + "SignatureType";
        public static readonly XName VerticalResolution = digsig + "VerticalResolution";
        public static readonly XName WindowsVersion = digsig + "WindowsVersion";
    }
}