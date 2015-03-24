using Microsoft.Crm.Sdk.Messages; 
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Cloudrocket.Crm.Plugins
{
    // Courtesy of Gonzalo Ruiz https://crm2011workflowutils.codeplex.com/
    internal class Helpers
    {
        private bool invalid = false;

        public static QueryExpression ConvertRollupQueryToQueryExpression(EntityReference rollupQueryEr, IOrganizationService service) {
            Entity rollupQuery = service.Retrieve(rollupQueryEr.LogicalName, rollupQueryEr.Id, new ColumnSet("fetchxml"));
            FetchXmlToQueryExpressionRequest req = new FetchXmlToQueryExpressionRequest() {
                FetchXml = rollupQuery["fetchxml"] as string,
            };

            FetchXmlToQueryExpressionResponse resp = (FetchXmlToQueryExpressionResponse)service.Execute(req);
            return resp.Query;
        }

        public static void Throw(string message, Exception innerException = null) {
            if (innerException == null)
            {
                throw new InvalidPluginExecutionException(message);
            }
            else
            {
                throw new InvalidPluginExecutionException(message, innerException);
            }
        }

        /// <summary>
        /// Gets the time zone by looking up the area code.
        /// </summary>
        /// <param name="telephone"></param>
        /// <returns></returns>
        public static string GetTimeZone(string telephone) {
            int? areaCode = new int();
            string digits = string.Empty;
            string timeZone = string.Empty;

            if (telephone != String.Empty)
            {
                digits = string.Join(null, Regex.Split(telephone, "[^\\d]"));
                digits = digits.StartsWith("1") ? digits.Substring(1, 3) : digits.Substring(0, 3);
                areaCode = Convert.ToInt32(digits);
            }

            return GetTimeZoneFromNaAreaCode(areaCode);
        }

        /// <summary>
        /// Checks that an email address is, in fact a correctly formatted email address. Note: Does
        /// not check that the address actually exists.
        /// </summary>
        /// <param name="strIn"></param>
        /// <returns></returns>
        public bool IsValidEmail(string strIn) {
            invalid = false;
            if (String.IsNullOrEmpty(strIn))
                return false;

            // Use IdnMapping class to convert Unicode domain names.
            try
            {
                strIn = Regex.Replace(strIn, @"(@)(.+)$", this.DomainMapper,
                                      RegexOptions.None, TimeSpan.FromMilliseconds(200));
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }

            if (invalid)
                return false;

            // Return true if strIn is in valid e-mail format.
            try
            {
                return Regex.IsMatch(strIn,
                      @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                      @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,24}))$",
                      RegexOptions.IgnoreCase
                      , TimeSpan.FromMilliseconds(250));
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
        }

        private string DomainMapper(Match match) {
            // IdnMapping class with default property values.
            IdnMapping idn = new IdnMapping();

            string domainName = match.Groups[2].Value;
            try
            {
                domainName = idn.GetAscii(domainName);
            }
            catch (ArgumentException)
            {
                invalid = true;
            }
            return match.Groups[1].Value + domainName;
        }

        /// <summary>
        /// Gets the time zone from a North American area code.
        /// </summary>
        /// <param name="areaCode"></param>
        /// <returns></returns>
        private static string GetTimeZoneFromNaAreaCode(int? areaCode) {
            string timeZone = String.Empty;

            switch (areaCode)
            {
                case 201: timeZone = "Eastern"; break;
                case 202: timeZone = "Eastern"; break;
                case 203: timeZone = "Eastern"; break;
                case 204: timeZone = "Central"; break;
                case 205: timeZone = "Central"; break;
                case 206: timeZone = "Pacific"; break;
                case 207: timeZone = "Eastern"; break;
                case 208: timeZone = "Mountain"; break;
                case 209: timeZone = "Pacific"; break;
                case 210: timeZone = "Central"; break;
                case 212: timeZone = "Eastern"; break;
                case 213: timeZone = "Pacific"; break;
                case 214: timeZone = "Central"; break;
                case 215: timeZone = "Eastern"; break;
                case 216: timeZone = "Eastern"; break;
                case 217: timeZone = "Central"; break;
                case 218: timeZone = "Central"; break;
                case 219: timeZone = "Central"; break;
                case 224: timeZone = "Central"; break;
                case 225: timeZone = "Central"; break;
                case 226: timeZone = "Eastern"; break;
                case 228: timeZone = "Central"; break;
                case 229: timeZone = "Eastern"; break;
                case 231: timeZone = "Eastern"; break;
                case 234: timeZone = "Eastern"; break;
                case 239: timeZone = "Eastern"; break;
                case 240: timeZone = "Eastern"; break;
                case 248: timeZone = "Eastern"; break;
                case 249: timeZone = "Eastern"; break;
                case 250: timeZone = "Pacific"; break;
                case 251: timeZone = "Central"; break;
                case 252: timeZone = "Eastern"; break;
                case 253: timeZone = "Pacific"; break;
                case 254: timeZone = "Central"; break;
                case 256: timeZone = "Central"; break;
                case 260: timeZone = "Eastern"; break;
                case 262: timeZone = "Central"; break;
                case 267: timeZone = "Eastern"; break;
                case 269: timeZone = "Eastern"; break;
                case 270: timeZone = "Central"; break;
                case 276: timeZone = "Eastern"; break;
                case 281: timeZone = "Central"; break;
                case 289: timeZone = "Eastern"; break;
                case 301: timeZone = "Eastern"; break;
                case 302: timeZone = "Eastern"; break;
                case 303: timeZone = "Mountain"; break;
                case 304: timeZone = "Eastern"; break;
                case 305: timeZone = "Eastern"; break;
                case 306: timeZone = "Mountain"; break;
                case 307: timeZone = "Mountain"; break;
                case 308: timeZone = "Central"; break;
                case 309: timeZone = "Central"; break;
                case 310: timeZone = "Pacific"; break;
                case 312: timeZone = "Central"; break;
                case 313: timeZone = "Eastern"; break;
                case 314: timeZone = "Central"; break;
                case 315: timeZone = "Eastern"; break;
                case 316: timeZone = "Central"; break;
                case 317: timeZone = "Eastern"; break;
                case 318: timeZone = "Central"; break;
                case 319: timeZone = "Central"; break;
                case 320: timeZone = "Central"; break;
                case 321: timeZone = "Eastern"; break;
                case 323: timeZone = "Pacific"; break;
                case 325: timeZone = "Central"; break;
                case 330: timeZone = "Eastern"; break;
                case 331: timeZone = "Central"; break;
                case 334: timeZone = "Central"; break;
                case 336: timeZone = "Eastern"; break;
                case 337: timeZone = "Central"; break;
                case 339: timeZone = "Eastern"; break;
                case 340: timeZone = "Atlantic"; break;
                case 343: timeZone = "Eastern"; break;
                case 347: timeZone = "Eastern"; break;
                case 351: timeZone = "Eastern"; break;
                case 352: timeZone = "Eastern"; break;
                case 360: timeZone = "Pacific"; break;
                case 361: timeZone = "Central"; break;
                case 385: timeZone = "Mountain"; break;
                case 386: timeZone = "Eastern"; break;
                case 401: timeZone = "Eastern"; break;
                case 402: timeZone = "Central"; break;
                case 403: timeZone = "Mountain"; break;
                case 404: timeZone = "Eastern"; break;
                case 405: timeZone = "Central"; break;
                case 406: timeZone = "Mountain"; break;
                case 407: timeZone = "Eastern"; break;
                case 408: timeZone = "Pacific"; break;
                case 409: timeZone = "Central"; break;
                case 410: timeZone = "Eastern"; break;
                case 412: timeZone = "Eastern"; break;
                case 413: timeZone = "Eastern"; break;
                case 414: timeZone = "Central"; break;
                case 415: timeZone = "Pacific"; break;
                case 416: timeZone = "Eastern"; break;
                case 417: timeZone = "Central"; break;
                case 418: timeZone = "Eastern"; break;
                case 419: timeZone = "Eastern"; break;
                case 423: timeZone = "Eastern"; break;
                case 424: timeZone = "Pacific"; break;
                case 425: timeZone = "Pacific"; break;
                case 430: timeZone = "Central"; break;
                case 432: timeZone = "Central"; break;
                case 434: timeZone = "Eastern"; break;
                case 435: timeZone = "Mountain"; break;
                case 438: timeZone = "Eastern"; break;
                case 440: timeZone = "Eastern"; break;
                case 442: timeZone = "Pacific"; break;
                case 443: timeZone = "Eastern"; break;
                case 450: timeZone = "Eastern"; break;
                case 458: timeZone = "Pacific"; break;
                case 469: timeZone = "Central"; break;
                case 470: timeZone = "Eastern"; break;
                case 475: timeZone = "Eastern"; break;
                case 478: timeZone = "Eastern"; break;
                case 479: timeZone = "Central"; break;
                case 480: timeZone = "Mountain"; break;
                case 484: timeZone = "Eastern"; break;
                case 501: timeZone = "Central"; break;
                case 502: timeZone = "Eastern"; break;
                case 503: timeZone = "Pacific"; break;
                case 504: timeZone = "Central"; break;
                case 505: timeZone = "Mountain"; break;
                case 506: timeZone = "Atlantic"; break;
                case 507: timeZone = "Central"; break;
                case 508: timeZone = "Eastern"; break;
                case 509: timeZone = "Pacific"; break;
                case 510: timeZone = "Pacific"; break;
                case 512: timeZone = "Central"; break;
                case 513: timeZone = "Eastern"; break;
                case 514: timeZone = "Eastern"; break;
                case 515: timeZone = "Central"; break;
                case 516: timeZone = "Eastern"; break;
                case 517: timeZone = "Eastern"; break;
                case 518: timeZone = "Eastern"; break;
                case 519: timeZone = "Eastern"; break;
                case 520: timeZone = "Mountain"; break;
                case 530: timeZone = "Pacific"; break;
                case 531: timeZone = "Central"; break;
                case 534: timeZone = "Central"; break;
                case 539: timeZone = "Central"; break;
                case 540: timeZone = "Eastern"; break;
                case 541: timeZone = "Pacific"; break;
                case 551: timeZone = "Eastern"; break;
                case 559: timeZone = "Pacific"; break;
                case 561: timeZone = "Eastern"; break;
                case 562: timeZone = "Pacific"; break;
                case 563: timeZone = "Central"; break;
                case 567: timeZone = "Eastern"; break;
                case 570: timeZone = "Eastern"; break;
                case 571: timeZone = "Eastern"; break;
                case 573: timeZone = "Central"; break;
                case 574: timeZone = "Eastern"; break;
                case 575: timeZone = "Mountain"; break;
                case 579: timeZone = "Eastern"; break;
                case 580: timeZone = "Central"; break;
                case 581: timeZone = "Eastern"; break;
                case 585: timeZone = "Eastern"; break;
                case 586: timeZone = "Eastern"; break;
                case 587: timeZone = "Mountain"; break;
                case 601: timeZone = "Central"; break;
                case 602: timeZone = "Mountain"; break;
                case 603: timeZone = "Eastern"; break;
                case 604: timeZone = "Pacific"; break;
                case 605: timeZone = "Central"; break;
                case 606: timeZone = "Eastern"; break;
                case 607: timeZone = "Eastern"; break;
                case 608: timeZone = "Central"; break;
                case 609: timeZone = "Eastern"; break;
                case 610: timeZone = "Eastern"; break;
                case 612: timeZone = "Central"; break;
                case 613: timeZone = "Eastern"; break;
                case 614: timeZone = "Eastern"; break;
                case 615: timeZone = "Central"; break;
                case 616: timeZone = "Eastern"; break;
                case 617: timeZone = "Eastern"; break;
                case 618: timeZone = "Central"; break;
                case 619: timeZone = "Pacific"; break;
                case 620: timeZone = "Central"; break;
                case 623: timeZone = "Mountain"; break;
                case 626: timeZone = "Pacific"; break;
                case 630: timeZone = "Central"; break;
                case 631: timeZone = "Eastern"; break;
                case 636: timeZone = "Central"; break;
                case 641: timeZone = "Central"; break;
                case 646: timeZone = "Eastern"; break;
                case 647: timeZone = "Eastern"; break;
                case 650: timeZone = "Pacific"; break;
                case 651: timeZone = "Central"; break;
                case 657: timeZone = "Pacific"; break;
                case 660: timeZone = "Central"; break;
                case 661: timeZone = "Pacific"; break;
                case 662: timeZone = "Central"; break;
                case 667: timeZone = "Eastern"; break;
                case 678: timeZone = "Eastern"; break;
                case 681: timeZone = "Eastern"; break;
                case 682: timeZone = "Central"; break;
                case 701: timeZone = "Central"; break;
                case 702: timeZone = "Pacific"; break;
                case 703: timeZone = "Eastern"; break;
                case 704: timeZone = "Eastern"; break;
                case 705: timeZone = "Eastern"; break;
                case 706: timeZone = "Eastern"; break;
                case 707: timeZone = "Pacific"; break;
                case 708: timeZone = "Central"; break;
                case 709: timeZone = "Atlantic"; break;
                case 712: timeZone = "Central"; break;
                case 713: timeZone = "Central"; break;
                case 714: timeZone = "Pacific"; break;
                case 715: timeZone = "Central"; break;
                case 716: timeZone = "Eastern"; break;
                case 717: timeZone = "Eastern"; break;
                case 718: timeZone = "Eastern"; break;
                case 719: timeZone = "Mountain"; break;
                case 720: timeZone = "Mountain"; break;
                case 724: timeZone = "Eastern"; break;
                case 727: timeZone = "Eastern"; break;
                case 731: timeZone = "Central"; break;
                case 732: timeZone = "Eastern"; break;
                case 734: timeZone = "Eastern"; break;
                case 740: timeZone = "Eastern"; break;
                case 747: timeZone = "Pacific"; break;
                case 754: timeZone = "Eastern"; break;
                case 757: timeZone = "Eastern"; break;
                case 760: timeZone = "Pacific"; break;
                case 762: timeZone = "Eastern"; break;
                case 763: timeZone = "Central"; break;
                case 765: timeZone = "Eastern"; break;
                case 769: timeZone = "Central"; break;
                case 770: timeZone = "Eastern"; break;
                case 772: timeZone = "Eastern"; break;
                case 773: timeZone = "Central"; break;
                case 774: timeZone = "Eastern"; break;
                case 775: timeZone = "Pacific"; break;
                case 778: timeZone = "Pacific"; break;
                case 779: timeZone = "Central"; break;
                case 780: timeZone = "Mountain"; break;
                case 781: timeZone = "Eastern"; break;
                case 785: timeZone = "Central"; break;
                case 786: timeZone = "Eastern"; break;
                case 801: timeZone = "Mountain"; break;
                case 802: timeZone = "Eastern"; break;
                case 803: timeZone = "Eastern"; break;
                case 804: timeZone = "Eastern"; break;
                case 805: timeZone = "Pacific"; break;
                case 806: timeZone = "Central"; break;
                case 807: timeZone = "Eastern"; break;
                case 808: timeZone = "Hawaii"; break;
                case 810: timeZone = "Eastern"; break;
                case 812: timeZone = "Eastern"; break;
                case 813: timeZone = "Eastern"; break;
                case 814: timeZone = "Eastern"; break;
                case 815: timeZone = "Central"; break;
                case 816: timeZone = "Central"; break;
                case 817: timeZone = "Central"; break;
                case 818: timeZone = "Pacific"; break;
                case 819: timeZone = "Eastern"; break;
                case 828: timeZone = "Eastern"; break;
                case 830: timeZone = "Central"; break;
                case 831: timeZone = "Pacific"; break;
                case 832: timeZone = "Central"; break;
                case 843: timeZone = "Eastern"; break;
                case 845: timeZone = "Eastern"; break;
                case 847: timeZone = "Central"; break;
                case 848: timeZone = "Eastern"; break;
                case 850: timeZone = "Central"; break;
                case 856: timeZone = "Eastern"; break;
                case 857: timeZone = "Eastern"; break;
                case 858: timeZone = "Pacific"; break;
                case 859: timeZone = "Eastern"; break;
                case 860: timeZone = "Eastern"; break;
                case 862: timeZone = "Eastern"; break;
                case 863: timeZone = "Eastern"; break;
                case 864: timeZone = "Eastern"; break;
                case 865: timeZone = "Eastern"; break;
                case 867: timeZone = "Pacific"; break;
                case 870: timeZone = "Central"; break;
                case 872: timeZone = "Central"; break;
                case 878: timeZone = "Eastern"; break;
                case 901: timeZone = "Central"; break;
                case 902: timeZone = "Atlantic"; break;
                case 903: timeZone = "Central"; break;
                case 904: timeZone = "Eastern"; break;
                case 905: timeZone = "Eastern"; break;
                case 906: timeZone = "Eastern"; break;
                case 907: timeZone = "Alaska"; break;
                case 908: timeZone = "Eastern"; break;
                case 909: timeZone = "Pacific"; break;
                case 910: timeZone = "Eastern"; break;
                case 912: timeZone = "Eastern"; break;
                case 913: timeZone = "Central"; break;
                case 914: timeZone = "Eastern"; break;
                case 915: timeZone = "Mountain"; break;
                case 916: timeZone = "Pacific"; break;
                case 917: timeZone = "Eastern"; break;
                case 918: timeZone = "Central"; break;
                case 919: timeZone = "Eastern"; break;
                case 920: timeZone = "Central"; break;
                case 925: timeZone = "Pacific"; break;
                case 928: timeZone = "Mountain"; break;
                case 929: timeZone = "Eastern"; break;
                case 931: timeZone = "Central"; break;
                case 936: timeZone = "Central"; break;
                case 937: timeZone = "Eastern"; break;
                case 938: timeZone = "Central"; break;
                case 940: timeZone = "Central"; break;
                case 941: timeZone = "Eastern"; break;
                case 947: timeZone = "Eastern"; break;
                case 949: timeZone = "Pacific"; break;
                case 951: timeZone = "Pacific"; break;
                case 952: timeZone = "Central"; break;
                case 954: timeZone = "Eastern"; break;
                case 956: timeZone = "Central"; break;
                case 970: timeZone = "Mountain"; break;
                case 971: timeZone = "Pacific"; break;
                case 972: timeZone = "Central"; break;
                case 973: timeZone = "Eastern"; break;
                case 978: timeZone = "Eastern"; break;
                case 979: timeZone = "Central"; break;
                case 980: timeZone = "Eastern"; break;
                case 984: timeZone = "Eastern"; break;
                case 985: timeZone = "Central"; break;
                case 989: timeZone = "Eastern"; break;

                default: timeZone = "Unknown"; break;
            }

            return timeZone;
        }

        /// <summary>
        /// Gets the ShowMe price per user in USD.
        /// </summary>
        public static decimal GetShowMePricePerUser(decimal users) {
            decimal pricePerUser = 10;
            decimal marginalUsers;

            if (users <= 150)
            {
                pricePerUser = (users * 10) / users;
                return Math.Round(pricePerUser, 2);
            }
            if (users <= 250)
            {
                marginalUsers = users - 150;
                pricePerUser = ((150 * 10) + (marginalUsers * 9)) / users;
                return Math.Round(pricePerUser, 2);
            }
            if (users <= 500)
            {
                marginalUsers = users - 250;
                pricePerUser = ((150 * 10) + (100 * 9) + (marginalUsers * 8)) / users;
                return Math.Round(pricePerUser, 2);
            }
            if (users <= 1000)
            {
                marginalUsers = users - 500;
                pricePerUser = ((150 * 10) + (100 * 9) + (250 * 8) + (marginalUsers * 7)) / users;
                return Math.Round(pricePerUser, 2);
            }
            if (users <= 2500)
            {
                marginalUsers = users - 1000;
                pricePerUser = ((150 * 10) + (100 * 9) + (250 * 8) + (500 * 7) + (marginalUsers * 6)) / users;
                return Math.Round(pricePerUser, 2);
            }
            if (users <= 5000)
            {
                marginalUsers = users - 2500;
                pricePerUser = ((150 * 10) + (100 * 9) + (250 * 8) + (500 * 7) + (1500 * 6) + (marginalUsers * 5)) / users;
                return Math.Round(pricePerUser, 2);
            }
            if (users > 5000)
            {
                pricePerUser = 29450 / users;
                return Math.Round(pricePerUser, 2);
            }

            return Math.Round(pricePerUser, 2);
        }
    }
}