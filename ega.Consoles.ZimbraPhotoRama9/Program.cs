using System;
using System.Text;
using System.Net;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using System.Security.Cryptography.X509Certificates;

namespace ega.Consoles.ZimbraPhotoRama9
{
    class Program : ICertificatePolicy
    {
        public bool CheckValidationResult(ServicePoint srvPoint,
        X509Certificate certificate, WebRequest request,
        int certificateProblem)
        {
            //Return True to force the certificate to be accepted.
            return true;
        }

        static void SendNotification(string argMsg)
        {
            var request = (HttpWebRequest)WebRequest.Create("https://notify-api.line.me/api/notify");
            var postData = string.Format("message={0}", argMsg);
            var data = Encoding.UTF8.GetBytes(postData);

            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = data.Length;
            request.Headers.Add("Authorization", "Bearer TRp6byyCsJG7S2poh5ON3zdH88SSm3LMffZ1fXy8o1H"); //KlPhgOKMqBYSuLsZBLAY7uUCXD1s0jEjwHfbUPbQE0I

            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            var response = (HttpWebResponse)request.GetResponse();
            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

            Console.WriteLine("\r\n{0} \r\n", responseString);
        }

        /// <summary>
        /// Send notication to EGA-CSD LINE Noti
        /// </summary>
        /// <param name="argtoken"></param>
        static void CheckAndSendLineNotification(string argtoken)
        {
            if (!string.IsNullOrEmpty(argtoken))
            {
                //Console.WriteLine("not need to send notification..");
                //SendNotification("[MailGoThai]: สามารถเชื่อมต่อ MailGoThai ได้ (เทสโดย : mailchecker@training.mail.go.th)");
            }
            else
            {
                SendNotification("[MailGoThai]: ไม่สามารถดึงค่า token ได้ โปรดติดต่อผู้ดูแลระบบ");
            }
        }

        static void DownloadRemoteImageFile(string uri, string fileName)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            //request.Credentials = new System.Net.NetworkCredential("mailchecker@training.mail.go.th", "Ega@2017");
            
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            
            if ((response.StatusCode == HttpStatusCode.OK ||
                response.StatusCode == HttpStatusCode.Moved ||
                response.StatusCode == HttpStatusCode.Redirect) &&
                response.ContentType.StartsWith("image", StringComparison.OrdinalIgnoreCase))
            {
                using(Stream inputStream = response.GetResponseStream())
                using (Stream outputStream = File.OpenWrite(fileName))
                {
                    byte[] buffer = new byte[4096];
                    int byteRead;
                    do
                    {
                        byteRead = inputStream.Read(buffer, 0, buffer.Length);
                        outputStream.Write(buffer, 0, byteRead);
                    } while (byteRead != 0);

                }

            }

        }

        static string GetMailByFolder(string argtoken,string argFolderQuery,string argOfficerEmail)
        {
            string sret = string.Empty;

            
            StringBuilder sbXML2 = new StringBuilder();
            sbXML2.Append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
            sbXML2.Append("<soap:Envelope xmlns:soap=\"http://www.w3.org/2003/05/soap-envelope\">");
            sbXML2.Append("<soap:Header>");
            sbXML2.Append("<context xmlns=\"urn:zimbra\">");
            sbXML2.Append("<format type=\"xml\"/>");
            sbXML2.Append(string.Format("<authToken>{0}</authToken>", argtoken));
            sbXML2.Append("</context>");
            sbXML2.Append("</soap:Header>");
            sbXML2.Append("<soap:Body>");
            sbXML2.Append("<SearchRequest limit=\"2000\" ");
            sbXML2.Append(string.Format(" xmlns=\"urn:zimbraMail\">"));
            sbXML2.Append(string.Format("<query>{0}</query>", argFolderQuery));
            sbXML2.Append("</SearchRequest>");
            sbXML2.Append("</soap:Body>");
            sbXML2.Append("</soap:Envelope>");
            Console.WriteLine(sbXML2);

            HttpWebRequest httpZimbraRequest2 = (HttpWebRequest)WebRequest.Create("https://accounts.mail.go.th/service/soap");
            httpZimbraRequest2.ContentType = "application/soapxml";

            byte[] byteArray2 = Encoding.UTF8.GetBytes(sbXML2.ToString());

            httpZimbraRequest2.Method = "POST";
            httpZimbraRequest2.ContentLength = byteArray2.Length;

            using (var stream = httpZimbraRequest2.GetRequestStream())
            {
                stream.Write(byteArray2, 0, byteArray2.Length);
            }
            var response2 = (HttpWebResponse)httpZimbraRequest2.GetResponse();

            var responseString2 = new StreamReader(response2.GetResponseStream()).ReadToEnd();
            Console.WriteLine(responseString2);
            
            
            
            //@@ reload to XML
            XmlDocument responseDoc2 = new XmlDocument();
            responseDoc2.LoadXml(responseString2);
            //XmlNodeList docs = responseDoc2.SelectNodes("/SearchResponse");
            XmlNodeList docs = responseDoc2.GetElementsByTagName("c");
            Console.WriteLine("have {0} nodes", docs.Count);
            for (int i = 0; i < docs.Count; i++)
            {
                if (docs[i].InnerXml.Length > 0)
                {
                    string sInnerXml = docs[i].InnerXml;
                    Console.WriteLine("{0}). {1}", i, sInnerXml);
                    string sId = GetIdInnerXml(sInnerXml);
                    Console.WriteLine("id = {0}", sId);
                    string sEmailFrom = GetEmailFromInnerXml(sInnerXml);
                    Console.WriteLine("from = {0}", sEmailFrom);

                    string sPath = string.Format(@".\{0}\{1}{2:yyyyMMddHH}id{3}",argOfficerEmail,sEmailFrom, System.DateTime.Now,sId) ;  //@@ include file name
                    try
                    {
                        if (!Directory.Exists(sPath))
                        {
                            Directory.CreateDirectory(sPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                    //if (sId == "257")
                    //{
                        Console.WriteLine("@@@@@@@@@@@@@@@@@@@@@@@@ next to call Get-Msg @@@@@@@@@@@@@@@@@@@@@");
                        string smsg = GetMsgById(argtoken, sId,sEmailFrom,sPath);

                    //}

                    //XmlNode node = docs[i].SelectNodes("//c");
                    //string val = node.SelectSingleNode("m").Attributes["id"].Value;
                    //Console.WriteLine("id value = {0}", val);
                }
            }
             
            
            return sret;
        }

        static string GetMsgById(string argtoken, string argId,string argEmailFrom,string argPath)
        {
            string sret = string.Empty;

            StringBuilder sbXML2 = new StringBuilder();
            sbXML2.Append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
            sbXML2.Append("<soap:Envelope xmlns:soap=\"http://www.w3.org/2003/05/soap-envelope\">");
            sbXML2.Append("<soap:Header>");
            sbXML2.Append("<context xmlns=\"urn:zimbra\">");
            sbXML2.Append("<format type=\"xml\"/>");
            sbXML2.Append(string.Format("<authToken>{0}</authToken>", argtoken));
            sbXML2.Append("</context>");
            sbXML2.Append("</soap:Header>");
            sbXML2.Append("<soap:Body>");
            sbXML2.Append("<GetMsgRequest ");
            sbXML2.Append(string.Format(" xmlns=\"urn:zimbraMail\">"));
            sbXML2.Append(string.Format("<m id=\"{0}\" />", argId));
            sbXML2.Append("</GetMsgRequest>");
            sbXML2.Append("</soap:Body>");
            sbXML2.Append("</soap:Envelope>");
            Console.WriteLine(sbXML2);

            HttpWebRequest httpZimbraRequest2 = (HttpWebRequest)WebRequest.Create("https://accounts.mail.go.th/service/soap");
            httpZimbraRequest2.ContentType = "application/soapxml";

            byte[] byteArray2 = Encoding.UTF8.GetBytes(sbXML2.ToString());

            httpZimbraRequest2.Method = "POST";
            httpZimbraRequest2.ContentLength = byteArray2.Length;

            using (var stream = httpZimbraRequest2.GetRequestStream())
            {
                stream.Write(byteArray2, 0, byteArray2.Length);
            }
            var response2 = (HttpWebResponse)httpZimbraRequest2.GetResponse();

            var responseString2 = new StreamReader(response2.GetResponseStream()).ReadToEnd();
            Console.WriteLine("<<<<<<<<<<<<<<<<<<<<<<< Response part >>>>>>>>>>>>>>>>>>");
            Console.WriteLine(responseString2);

            XmlDocument responseDoc = new XmlDocument();
            responseDoc.LoadXml(responseString2);
            XmlNodeList docs = responseDoc.GetElementsByTagName("content");
            string scontent = string.Empty;
            if (docs[0].InnerText.Length > 0)
            {
                scontent = docs[0].InnerText;
            }
            Console.WriteLine("content ={0}", scontent);

            Console.WriteLine("Catch by Pattern below.........");
            string sPattern = "<mp cd=\"attachment\"(.*?)/>";
            MatchCollection mtResults = Regex.Matches(responseString2, sPattern);
            bool bHaveNewItem = false;
            string sMailBody = string.Empty;
            int iCount = 0;
            StreamWriter writerLinkResult = new StreamWriter( string.Format(@"{0}\{1}id{2}_link.html", argPath,argEmailFrom,argId), true);
            writerLinkResult.AutoFlush = true;
            StreamWriter writerContent = new StreamWriter(string.Format(@"{0}\{1}id{2}_content.html", argPath, argEmailFrom, argId), true);
            writerContent.AutoFlush = true;

            writerContent.WriteLine(scontent);
            writerLinkResult.WriteLine("Image links for this {0} email",argEmailFrom);
            writerLinkResult.WriteLine("</br></br>");
            foreach (Match m in mtResults)
            {
                iCount +=1;
                string sValue = m.Value;
                Console.WriteLine(sValue);
                
                string sFilename = sValue.Substring(sValue.IndexOf("filename=\"") + 10, (sValue.IndexOf("s=\"") - sValue.IndexOf("filename=\""))-12);
                string sPart = sValue.Substring(sValue.IndexOf("part=\"") + 6, (sValue.IndexOf("\"/>") - sValue.IndexOf("part=\"")) - 6);
                Console.WriteLine("filename ={0}", sFilename);
                Console.WriteLine("Part ={0}", sPart);
                string simglink = string.Format(@"https://accounts.mail.go.th/service/home/~/?auth=co&loc=th&id={0}&part={1}",argId,sPart);
                downloadImage(simglink + "&disp=a", sFilename, argPath, argtoken);
                string sHref = string.Format("{0}.) ", iCount) + string.Format("<a href=\"{0}\">{1}</a>",simglink,sFilename) + Environment.NewLine + "</br>" + string.Format("\r\n");
                writerLinkResult.WriteLine(sHref);
            }

            writerLinkResult.Close();
            writerContent.Close();
            return sret;
        }

        /// <summary>
        /// Get Id Mail MSG
        /// </summary>
        /// <param name="argXml"></param>
        /// <returns></returns>
        static string GetIdInnerXml(string argXml)
        {
            string ret = string.Empty;

            if (argXml.Length > 0)
            {
                ret = argXml.Substring(argXml.IndexOf("id=\"") + 4, (argXml.IndexOf("l=\"") - argXml.IndexOf("id=\"")) - 6);

            }

            return ret;
        }

        static string GetEmailFromInnerXml(string argXml)
        {
            string ret = string.Empty;

            if (argXml.Length > 0)
            {
                ret = argXml.Substring(argXml.IndexOf("a=\"") + 3, (argXml.IndexOf("d=\"") - argXml.IndexOf("a=\"")) - 5);
            }

            return ret;
        }

        public static void downloadImage(String imgNameURL, String desImgName, String desPath, String authToken)
        {
            //example : string imageUrl = @"https://accounts.mail.go.th/service/home/~/?auth=co&loc=th&id=546&part=2&disp=a";
            string imageUrl = imgNameURL;

            //string saveLocation = @desPath+desImgName;
            string saveLocation = desPath +"\\"+ desImgName;
            Console.Write("url: {0} to {1}", imageUrl, saveLocation);
            /*
             * ZM_TEST=true; ZM_AUTH_TOKEN=0_507b84560d769e763b1047f48173f7d150628e73_69643d33363a31363161333161322d356561662d346461362d626139342d3439653335333464393361633b6578703d31333a313530383536303135373733363b76763d313a303b747970653d363a7a696d6272613b7469643d31303a313733353031313834393b76657273696f6e3d31333a382e362e305f47415f313135333b637372663d313a313b; JSESSIONID=bpuo4me8cm875llok04lw64k
             */

            byte[] imageBytes;
            HttpWebRequest imageRequest = (HttpWebRequest)WebRequest.Create(imageUrl);
            imageRequest.Headers.Add("Cookie", "ZM_AUTH_TOKEN=" + authToken);

            WebResponse imageResponse = imageRequest.GetResponse();

            Stream responseStream = imageResponse.GetResponseStream();

            using (BinaryReader br = new BinaryReader(responseStream))
            {
                imageBytes = br.ReadBytes(50000000);
                br.Close();
            }
            responseStream.Close();
            imageResponse.Close();

            FileStream fs = new FileStream(saveLocation, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            try
            {
                bw.Write(imageBytes);
            }
            finally
            {
                fs.Close();
                bw.Close();
            }
        }
        static void Main(string[] args)
        {

            string sFolderQuery = "in:inbox";
            string sOfficerEmail = "mailchecker@training.mail.go.th";
            string sOfficerEmailPass = "...";
            if (args != null)
            {
                ///Console.WriteLine("argument length = {0}", args.Length);
                if (args.Length < 4)
                {
                    if (args.Length == 1)
                    {
                        sFolderQuery += string.Format("/{0}", args[0]);
                    }
                    else if (args.Length > 1 && args.Length < 4)
                    {
                        sFolderQuery += string.Format("/{0}", args[0]);
                        sOfficerEmail = args[1];
                        sOfficerEmailPass = args[2];
                    }
                }
                else if(args.Length >= 4 )
                {
                    Console.WriteLine("This version can use only 1 sub level folder (from inbox:) \r\n Usage: ega.Consoles.ZimbraPhotoRama9.exe <<Folder Name>>");
                    Console.WriteLine("For exammple: \r\n ega.Consoles.ZimbraPhotoRama9.exe Bangkok p1@mail.go.th Ph@toking91 \r\n (It will get all emails under inbox/Bangkok)");
                    return;
                }
            }
            
            string stoken = string.Empty;
            ServicePointManager.CertificatePolicy = new Program();
            StringBuilder xml = new StringBuilder();
            xml.Append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
            xml.Append("<soap:Envelope xmlns:soap=\"http://www.w3.org/2003/05/soap-envelope\">");
            xml.Append("<soap:Header>");
            xml.Append("<context xmlns=\"urn:zimbra\">");
            xml.Append("<format type=\"xml\"/>");
            xml.Append("</context>");
            xml.Append("</soap:Header>");
            xml.Append("<soap:Body>");
            xml.Append("<AuthRequest xmlns=\"urn:zimbraAccount\">");
            xml.Append(string.Format("<account by=\"name\">{0}</account>",sOfficerEmail));    
            xml.Append(string.Format("<password>{0}</password>",sOfficerEmailPass));                                      
            xml.Append("</AuthRequest>");
            xml.Append("</soap:Body>");
            xml.Append("</soap:Envelope>");
            Console.WriteLine(xml);
            HttpWebRequest httpRequest = (HttpWebRequest)WebRequest.Create("https://accounts.mail.go.th/service/soap");
            httpRequest.ContentType = "application/soapxml";

            byte[] byteArray = Encoding.UTF8.GetBytes(xml.ToString());

            httpRequest.Method = "POST";
            httpRequest.ContentLength = byteArray.Length;
            try
            {
                using (var stream = httpRequest.GetRequestStream())
                {
                    stream.Write(byteArray, 0, byteArray.Length);
                }
                var response = (HttpWebResponse)httpRequest.GetResponse();

                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                Console.WriteLine(responseString);

                //Console.ReadLine();

                XmlDocument responseDoc = new XmlDocument();
                responseDoc.LoadXml(responseString);
                string authtoken = responseDoc.GetElementsByTagName("authToken").Item(0).InnerXml;
                stoken = authtoken;
                Console.WriteLine("\r\nauthToken:{0}\r\n", stoken);
                
                //CheckAndSendLineNotification(stoken);
                

                //@@ next call other Zimbra API function here.
                

                 string sMails = string.Empty;  //might be use List<string> instead.
                 sMails = GetMailByFolder(stoken, sFolderQuery,sOfficerEmail); // if want to sub folder for example "Bangkok" use , "in:inbox/Bangkok"

                //DownloadRemoteImageFile(@"https://accounts.mail.go.th/service/home/~/?auth=co&loc=th&id=262&part=2.4", "testImage001.jpg");
                //byte[] data;
                //using (WebClient client = new WebClient())
                //{
                //    client.UseDefaultCredentials = true;
                //    client.Credentials = new NetworkCredential("mailchecker@training.mail.go.th", "Ega@2017");
                //    data = client.DownloadData("https://accounts.mail.go.th/service/home/~/?auth=co&loc=th&id=262&part=2.4");
                //}
                //File.WriteAllBytes("testimage.jpg", data);
            }
            catch (WebException wex)
            {
                var pageContent = new StreamReader(wex.Response.GetResponseStream()).ReadToEnd();
                Console.WriteLine("Web Error: \r\n{0}", pageContent);
                //Console.WriteLine("send notification to LINE..");
                //SendNotification("[MailGoThai]: ไม่สามารถเชื่อมต่อ MailGoThai ได้ โปรดติดต่อผู้ดูแลระบบ");
            }
        }
    }
}
