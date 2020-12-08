using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class W3DIGSIG
    {
        public static readonly XNamespace w3digsig =
            "http://www.w3.org/2000/09/xmldsig#";

        public static readonly XName CanonicalizationMethod = w3digsig + "CanonicalizationMethod";
        public static readonly XName DigestMethod = w3digsig + "DigestMethod";
        public static readonly XName DigestValue = w3digsig + "DigestValue";
        public static readonly XName Exponent = w3digsig + "Exponent";
        public static readonly XName KeyInfo = w3digsig + "KeyInfo";
        public static readonly XName KeyValue = w3digsig + "KeyValue";
        public static readonly XName Manifest = w3digsig + "Manifest";
        public static readonly XName Modulus = w3digsig + "Modulus";
        public static readonly XName Object = w3digsig + "Object";
        public static readonly XName Reference = w3digsig + "Reference";
        public static readonly XName RSAKeyValue = w3digsig + "RSAKeyValue";
        public static readonly XName Signature = w3digsig + "Signature";
        public static readonly XName SignatureMethod = w3digsig + "SignatureMethod";
        public static readonly XName SignatureProperties = w3digsig + "SignatureProperties";
        public static readonly XName SignatureProperty = w3digsig + "SignatureProperty";
        public static readonly XName SignatureValue = w3digsig + "SignatureValue";
        public static readonly XName SignedInfo = w3digsig + "SignedInfo";
        public static readonly XName Transform = w3digsig + "Transform";
        public static readonly XName Transforms = w3digsig + "Transforms";
        public static readonly XName X509Certificate = w3digsig + "X509Certificate";
        public static readonly XName X509Data = w3digsig + "X509Data";
        public static readonly XName X509IssuerName = w3digsig + "X509IssuerName";
        public static readonly XName X509SerialNumber = w3digsig + "X509SerialNumber";
    }
}