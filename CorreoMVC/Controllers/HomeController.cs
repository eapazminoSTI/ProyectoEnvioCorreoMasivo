using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace CorreoMVC.Controllers
{
    public class HomeController : Controller
    {
        static SmtpClient SmtpClient { get; set; }
        List<String> listaTag = new List<String>();
        List<String> listaTagCol = new List<String>();
        public ActionResult Index()
        {
            string path1 = string.Format("{0}", Server.MapPath("~/Content/Uploads/Plantillas"));

            string[] ubicacion = Directory.GetFiles(path1);
            string[] archivos = new string[ubicacion.Length];
            for (int i = 0; i < ubicacion.Length; i++)
            {
                archivos[i] = (Path.GetFileName(ubicacion[i]));
            }
            ViewData["archivos"] = archivos;

            string[] columnasExcel = { "Por favor suba un archivo excel" };
            ViewData["colsExcel"] = columnasExcel;

            string[] emailExcel = { "Por favor suba un archivo excel" };
            ViewData["colsExcelEmail"] = emailExcel;


            return View();
        }

        public double Barra()
        {
            double numero = Convert.ToDouble(TempData["NumBarra"]);
            double resultado = 100.00 / (numero - 1);
            return (resultado);
        }

        public ActionResult Validar(string emailPropierty, string passwordPropierty, string smtpServer, string port)
        {
            try
            {
                SmtpClient = new SmtpClient
                {
                    Host = smtpServer,
                    Port = Convert.ToInt16(port),
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(emailPropierty, passwordPropierty)
                };
                var msg = new MailMessage
                {
                    From = new MailAddress(emailPropierty),
                    IsBodyHtml = true,
                    Body = "Su cuenta ha sido validada correctamente, ahora puede envíar correos.",

                    Subject = "Validación de cuenta correcta"
                };
                msg.To.Add(emailPropierty);
                SmtpClient.Send(msg);

                return Json(new { success = true, message = "Validación de Cuenta Correcta" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                string Error;
                Error = ex.Message;

                return Json(new { success = false, message = "Ha ocurrido un error al validar la cuenta: " + Error }, JsonRequestBehavior.AllowGet);

            }
        }

        public ActionResult Enviar(string emailPropierty, string passwordPropierty, string fromName, string xlsxFile, string subject, string emailHtmlTemplate, string smtpServer, string port, string fileAdjunto, string nameAdjunto, string colEx)
        {
            if (fromName != null | xlsxFile != null | subject != null | emailHtmlTemplate != null)
            {
                try
                {
                    xlsxFile = Convert.ToString(TempData["RutaExcel"]);
                    emailHtmlTemplate = Convert.ToString(TempData["Plantilla"]);
                    fileAdjunto = Convert.ToString(TempData["Adjunto"]);
                    nameAdjunto = Convert.ToString(TempData["NombreAdjunto"]);

                    string[] values = TempData["Asignaciones"] as string[]; //0:[nombre],nombre

                    int numero = values.Length;

                    string[] separadas = new string[2 * numero];
                    string[] nomTag = new string[numero];
                    string[] nomCol = new string[numero];

                    for (int i = 0; i < numero; i++)
                    {
                        if (separadas[0] == null)
                        {
                            separadas = values[i].Split(',').ToArray();
                        }
                        else
                        {
                            separadas = separadas.Concat(values[i].Split(',')).ToArray();
                        }

                    }
                    int aux = 0;
                    int aux1 = 0;
                    for (int j = 0; j < separadas.Length; j++)
                    {

                        if (j == 0 || j % 2 == 0)
                        {
                            nomTag[aux] = separadas[j];
                            aux++;
                        }
                        else
                        {
                            nomCol[aux1] = separadas[j];
                            aux1++;
                        }
                    }

                    SmtpClient = new SmtpClient
                    {
                        Host = smtpServer,
                        Port = Convert.ToInt16(port),
                        EnableSsl = true,
                        DeliveryMethod = SmtpDeliveryMethod.Network,
                        UseDefaultCredentials = false,
                        Credentials = new NetworkCredential(emailPropierty, passwordPropierty)
                    };


                    int delayBetweenEmails = 2000;

                    int startRow = 1;
                    if (startRow < 2)
                        startRow = 2;


                    var excelFile = new FileInfo(xlsxFile);
                    if (!excelFile.Exists)
                        throw new FileNotFoundException($"The file '{excelFile.FullName}' does not exist.", excelFile.FullName);


                    using (var package = new ExcelPackage(excelFile))
                    {
                        ExcelWorksheet workSheet = package.Workbook.Worksheets[0];

                        int mailsLeft = workSheet.Dimension.End.Row - startRow;
                        TimeSpan timeLeft = TimeSpan.FromMilliseconds(
                            mailsLeft * (delayBetweenEmails + 100));
                        Console.WriteLine("Estimated time: {0} hours", timeLeft);

                        // Read Excel file column titles (the first worhsheet row)
                        var columnIndexByName = new Dictionary<string, int>();
                        for (int col = 1; col <= workSheet.Dimension.End.Column; col++)
                        {
                            string columnName = workSheet.Cells[1, col].Value
                                .ToString().ToLower();
                            columnIndexByName[columnName] = col;
                        }
                        int dimentionRow = workSheet.Dimension.End.Row;
                        // Process all rows from the Excel file --> send email for each row

                        List<string> listaMensajes = new List<string>();

                        for (int row = startRow; row <= dimentionRow; row++)
                        {

                            string email = (string)workSheet.Cells[row, columnIndexByName[colEx]].Value; //nombre de la columna que contiene los emails

                            if ((email != null))
                            {
                                string bodyHtml = System.IO.File.ReadAllText(emailHtmlTemplate, Encoding.UTF8); //plantilla html

                                for (int i = 0; i < nomTag.Length; i++)
                                {
                                    string variable = "";
                                    variable = Convert.ToString(workSheet.Cells[row, columnIndexByName[nomCol[i].ToString()]].Value);

                                    bodyHtml = bodyHtml.Replace(nomTag[i], variable);
                                    //bodyHtml = bodyHtml.Replace("[img]", fileImagen);
                                }

                                var msg = new MailMessage
                                {
                                    From = new MailAddress(email, fromName),
                                    IsBodyHtml = true,
                                    Body = bodyHtml,
                                    Subject = subject,
                                };


                                if (fileAdjunto != "")
                                {
                                    FileStream fs = new FileStream(fileAdjunto, FileMode.Open, FileAccess.Read);
                                    string contentType = MimeMapping.GetMimeMapping(fileAdjunto);

                                    Attachment adjunto;
                                    adjunto = new Attachment(fs, nameAdjunto, contentType);

                                    msg.Attachments.Add(adjunto);
                                }

                                msg.To.Add(email);
                                SmtpClient.Send(msg);

                                Thread.Sleep(delayBetweenEmails);
                            }
                            else
                            {
                                listaMensajes.Add("No existe correo en la fila " + row);
                            }

                        }

                        return Json(new { success = true, message = "Correos enviados existosamente a los destinatarios", data = listaMensajes }, JsonRequestBehavior.AllowGet);
                    }
                }
                catch (Exception ex)
                {
                    string Error;
                    Error = ex.Message;

                    return Json(new { success = false, message = "Correos enviados, por favor revise el archivo de destinatarios: " + Error }, JsonRequestBehavior.AllowGet);

                }

            }

            return View("Index");
        }

        public ActionResult EnviarPlan(string emailPropierty, string passwordPropierty, string fromName, string xlsxFile, string subject, string emailHtmlTemplate, string smtpServer, string port, string fileAdjunto, string nameAdjunto, string fechaHora, string colEx)
        {
            double tiempoEnvio = 0;

            DateTime fechaActual = DateTime.Now;
            DateTime fechaUsuario = Convert.ToDateTime(fechaHora);

            if (fromName != null | xlsxFile != null | subject != null | emailHtmlTemplate != null)
            {
                if (fechaUsuario > fechaActual)
                {
                    try
                    {
                        //tiempoEnvio = (fechaUsuario - fechaActual).TotalMilliseconds;

                        //Thread.Sleep(Convert.ToInt32(tiempoEnvio));


                        new Thread(() =>
                        {
                            Thread.CurrentThread.IsBackground = true;

                            tiempoEnvio = (fechaUsuario - fechaActual).TotalMilliseconds;

                            Thread.Sleep(Convert.ToInt32(tiempoEnvio));


                            xlsxFile = Convert.ToString(TempData["RutaExcel"]);
                            emailHtmlTemplate = Convert.ToString(TempData["Plantilla"]);
                            fileAdjunto = Convert.ToString(TempData["Adjunto"]);
                            nameAdjunto = Convert.ToString(TempData["NombreAdjunto"]);
                            string[] values = TempData["Asignaciones"] as string[]; //0:[nombre],nombre

                            int numero = values.Length;

                            string[] separadas = new string[2 * numero];
                            string[] nomTag = new string[numero];
                            string[] nomCol = new string[numero];

                            for (int i = 0; i < numero; i++)
                            {
                                if (separadas[0] == null)
                                {
                                    separadas = values[i].Split(',').ToArray();
                                }
                                else
                                {
                                    separadas = separadas.Concat(values[i].Split(',')).ToArray();
                                }

                            }
                            int aux = 0;
                            int aux1 = 0;
                            for (int j = 0; j < separadas.Length; j++)
                            {

                                if (j == 0 || j % 2 == 0)
                                {
                                    nomTag[aux] = separadas[j];
                                    aux++;
                                }
                                else
                                {
                                    nomCol[aux1] = separadas[j];
                                    aux1++;
                                }
                            }

                            SmtpClient = new SmtpClient
                            {
                                Host = smtpServer,
                                Port = Convert.ToInt16(port),
                                EnableSsl = true,
                                DeliveryMethod = SmtpDeliveryMethod.Network,
                                UseDefaultCredentials = false,
                                Credentials = new NetworkCredential(emailPropierty, passwordPropierty)
                            };

                            int delayBetweenEmails = 2000;

                            int startRow = 1;
                            if (startRow < 2)
                                startRow = 2;


                            var excelFile = new FileInfo(xlsxFile);
                            if (!excelFile.Exists)
                                throw new FileNotFoundException($"The file '{excelFile.FullName}' does not exist.", excelFile.FullName);


                            using (var package = new ExcelPackage(excelFile))
                            {
                                ExcelWorksheet workSheet = package.Workbook.Worksheets[0];

                                int mailsLeft = workSheet.Dimension.End.Row - startRow;
                                TimeSpan timeLeft = TimeSpan.FromMilliseconds(
                                    mailsLeft * (delayBetweenEmails + 100));
                                Console.WriteLine("Estimated time: {0} hours", timeLeft);

                                // Read Excel file column titles (the first worhsheet row)
                                var columnIndexByName = new Dictionary<string, int>();
                                for (int col = 1; col <= workSheet.Dimension.End.Column; col++)
                                {
                                    string columnName = workSheet.Cells[1, col].Value
                                        .ToString().ToLower();
                                    columnIndexByName[columnName] = col;
                                }

                                // Process all rows from the Excel file --> send email for each row
                                for (int row = startRow; row <= workSheet.Dimension.End.Row; row++)
                                {
                                    string email = (string)workSheet.Cells[row, columnIndexByName[colEx]].Value; //nombre de la columna que contiene los emails


                                    string bodyHtml = System.IO.File.ReadAllText(emailHtmlTemplate, Encoding.UTF8); //plantilla html

                                    for (int i = 0; i < nomTag.Length; i++)
                                    {
                                        string variable = "";
                                        variable = Convert.ToString(workSheet.Cells[row, columnIndexByName[nomCol[i].ToString()]].Value);

                                        bodyHtml = bodyHtml.Replace(nomTag[i], variable);
                                    }

                                    var msg = new MailMessage
                                    {
                                        From = new MailAddress(email, fromName),
                                        IsBodyHtml = true,
                                        Body = bodyHtml,
                                        Subject = subject,
                                    };


                                    if (fileAdjunto != "")
                                    {
                                        FileStream fs = new FileStream(fileAdjunto, FileMode.Open, FileAccess.Read);
                                        string contentType = MimeMapping.GetMimeMapping(fileAdjunto);

                                        Attachment adjunto;
                                        adjunto = new Attachment(fs, nameAdjunto, contentType);

                                        msg.Attachments.Add(adjunto);
                                    }

                                    msg.To.Add(email);
                                    SmtpClient.Send(msg);

                                    Thread.Sleep(delayBetweenEmails);
                                }

                                Console.WriteLine("Hey, I'm from background thread");
                            }


                        }).Start();

                        return Json(new { success = true, message = "Correos planificados enviados existosamente a los destinatarios" }, JsonRequestBehavior.AllowGet);



                    }
                    catch (Exception ex)
                    {
                        string Error;
                        Error = ex.Message;

                        return Json(new { success = false, message = "Correos enviados, por favor revise el archivo de destinatarios: " + Error }, JsonRequestBehavior.AllowGet);
                    }

                }
                else
                {
                    return Json(new { success = false, message = "La fecha seleccionada debe ser mayor a la actual" }, JsonRequestBehavior.AllowGet);
                }
            }

            return View("Index");

        }
        public String separador(string cadena)
        {
            string resultado = "";
            char delimitador = '\\';
            string[] valores = cadena.Split(delimitador);
            for (int i = 0; i < valores.Length; i++)
            {
                if (i == 0)
                {
                    resultado = resultado + valores[i];
                }
                else
                {
                    resultado = resultado + "/" + valores[i];
                }

            }
            return resultado;
        }

        public ActionResult UploadFiles()
        {
            string path = "";
            string path1 = "";
            // Checking no of files injected in Request object  
            if (Request.Files.Count > 0)
            {
                try
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;
                    for (int i = 0; i < files.Count; i++)
                    {

                        HttpPostedFileBase file = files[i];
                        string fname;

                        // Checking for Internet Explorer  
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            fname = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            fname = file.FileName;
                        }

                        // Get the complete folder path and store the file inside it.
                        path = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads"), fname);

                        path1 = separador(path);

                        file.SaveAs(path1);

                    }

                    TempData["Ruta"] = path1;
                    TempData["Barra"] = path1;
                    TempData["RutaExcel"] = path1;
                    string xlsxFile = Convert.ToString(TempData["Barra"]);
                    var excelFile = new FileInfo(xlsxFile);
                    using (var package = new ExcelPackage(excelFile))
                    {
                        ExcelWorksheet workSheet = package.Workbook.Worksheets[0];
                        TempData["NumBarra"] = workSheet.Dimension.End.Row;

                    }

                    return Json(new { success = true, message = "Destinatarios subidos correctamente", path1 = path1 }, JsonRequestBehavior.AllowGet);

                }
                catch (Exception ex)
                {
                    return Json("Error occurred. Error details: " + ex.Message);
                }
            }
            else
            {
                return Json(new { success = false, message = "Primero debe seleccionar un archivo tipo excel" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult UploadFiles2(string nameFile)
        {
            string path = "";
            string path1 = "";
            string nombre = Path.GetFileNameWithoutExtension(nameFile);
            string pathTxt = "";

            try
            {
                // Get the complete folder path and store the file inside it.
                path = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads/Plantillas"), nameFile);
                pathTxt = string.Format("{0}/{1}", Server.MapPath("~/Content/Asignaciones"), nombre + ".txt");
                path1 = separador(path);


                TempData["Plantilla"] = path1;
                string[] asignaciones = System.IO.File.ReadAllLines(pathTxt);
                TempData["Asignaciones"] = asignaciones;

                //return Json(path1);
                return Json(new { success = true, message = "Plantilla subida correctamente", path1 = path1 }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                string Error;
                Error = ex.Message;

                return Json(new { success = false, message = "Ha ocurrido un error al subir la plantilla: " + Error }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult UploadFiles3()
        {
            string path = "";
            string path1 = "";
            string fileName;

            // Checking no of files injected in Request object  
            if (Request.Files.Count > 0)
            {
                try
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;
                    for (int i = 0; i < files.Count; i++)
                    {

                        HttpPostedFileBase file = files[i];
                        string fname;

                        // Checking for Internet Explorer  
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            fname = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            fname = file.FileName;
                        }

                        // Get the complete folder path and store the file inside it.
                        path = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads"), fname);
                        path1 = separador(path);

                        file.SaveAs(path1);

                    }

                    TempData["Adjunto"] = path1;
                    fileName = Path.GetFileName(path1);
                    TempData["NombreAdjunto"] = fileName;

                    //return Json(path1);
                    return Json(new { success = true, message = "Archivo subido correctamente", path1 = path1 }, JsonRequestBehavior.AllowGet);

                }
                catch (Exception ex)
                {
                    string Error;
                    Error = ex.Message;

                    return Json(new { success = false, message = "Ha ocurrido un error al subir el archivo: " + Error }, JsonRequestBehavior.AllowGet);
                }
            }
            else
            {
                return Json(new { success = false, message = "Primero debe seleccionar un archivo" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult UploadFiles4()
        {
            string path = "";
            string path1 = "";
            string fileName;

            // Checking no of files injected in Request object  
            if (Request.Files.Count > 0)
            {
                try
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;
                    for (int i = 0; i < files.Count; i++)
                    {

                        HttpPostedFileBase file = files[i];
                        string fname;

                        // Checking for Internet Explorer  
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            fname = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            fname = file.FileName;
                        }

                        // Get the complete folder path and store the file inside it.
                        path = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads/Imagenes"), fname);
                        path1 = separador(path);

                        file.SaveAs(path1);

                    }

                    TempData["Imagen"] = path1;
                    fileName = Path.GetFileName(path1);
                    TempData["NombreImagen"] = fileName;

                    return Json(new { success = true, message = "Imagen subida correctamente", path1 = path1 }, JsonRequestBehavior.AllowGet);

                }
                catch (Exception ex)
                {
                    string Error;
                    Error = ex.Message;

                    return Json(new { success = false, message = "Ha ocurrido un error al subir la imagen: " + Error }, JsonRequestBehavior.AllowGet);
                }
            }
            else
            {
                return Json(new { success = false, message = "Primero debe seleccionar una imagen" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult GrabarPlantillaPersonalizada(string nombrePlantilla, string tituloPlantilla, string linkSitio, string contenido, string firma, string imagen, string nombreImagen, List<String> values)
        {
            try
            {
                imagen = Convert.ToString(TempData["Imagen"]);
                nombreImagen = Convert.ToString(TempData["NombreImagen"]);

                TextWriter archivo;
                string path = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads/Plantillas"), nombrePlantilla);
                string path1 = string.Format("{0}", Server.MapPath("~/Content/Uploads/Plantillas"));
                string pathImg1 = string.Format("{0}/{1}", "https://www.socialhub-ec.com/correomvc/Content/Uploads/Imagenes", nombreImagen);

                archivo = new StreamWriter(path + ".html");

                string btnLinkSitio;

                if (linkSitio != "")
                {
                    btnLinkSitio = "<a href= https://" + linkSitio + " target='_blank'>Clic para conocer más</a>";
                }
                else
                {
                    btnLinkSitio = "";
                }

                string header = "<!doctype html><html>  <head>    <meta name='viewport' content='width=device-width' />    <meta http-equiv='Content-Type' content='text/html; charset=UTF-8' /><title> " + tituloPlantilla + " </title><style>img {border: none;-ms-interpolation-mode: bicubic;max-width: 100%;}body {background-color: #f6f6f6; font-family: sans-serif; -webkit-font-smoothing: antialiased;font-size: 14px;line-height: 1.4;margin: 0;padding: 0;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;}table {border-collapse: separate;mso-table-lspace: 0pt;mso-table-rspace: 0pt;width: 100%; }table td {font-family: sans-serif;font-size: 14px;vertical-align: top;}.body {background-color: #f6f6f6;width: 100%;} .container {display: block;margin: 0 auto !important; max-width: 580px;padding: 10px;width: 580px;}.content {box-sizing: border-box;display: block;margin: 0 auto;max-width: 580px;padding: 10px;}.main {background: #ffffff;border-radius: 3px;width: 100%;}.wrapper {box-sizing: border-box;padding: 20px;}.content-block {padding-bottom: 10px;padding-top: 10px;}.footer {clear: both;margin-top: 10px;text-align: center;width: 100%;}.footer td,.footer p,.footer span,.footer a {color: #999999;font-size: 12px;text-align: center;}h1,h2,h3,h4 {color: #000000;font-family: sans-serif;font-weight: 400;line-height: 1.4;margin: 0;margin-bottom: 30px;}h1 {font-size: 35px;font-weight: 300;text-align: center;text-transform: capitalize;}p,ul,ol {font-family: sans-serif;font-size: 14px;font-weight: normal;margin: 0; margin-bottom: 15px; } p li, ul li, ol li {list-style-position: inside;margin-left: 5px; } a {color: #3498db;text-decoration: underline;}.btn {box-sizing: border-box;width: 100%; }.btn > tbody > tr > td {padding-bottom: 15px; }.btn table {width: auto;}.btn table td {background-color: #ffffff;border-radius: 5px;text-align: center;}.btn a {background-color: #ffffff; border: solid 1px #3498db;border-radius: 5px;box-sizing: border-box;color: #3498db;cursor: pointer; display: inline-block;font-size: 14px;font-weight: bold; margin: 0; padding: 12px 25px; text-decoration: none; text-transform: capitalize;}.btn-primary table td { background-color: #3498db;} .btn-primary a { background-color: #3498db; border-color: #3498db; color: #ffffff;}.last {margin-bottom: 0;}.first{margin-top: 0;} .align-center {text-align: center;} .align-right {text-align: right;}.align-left {text-align: left;} .clear {clear: both;}.mt0 {margin-top: 0;}.mb0 {margin-bottom: 0;}.preheader {color: transparent;display: none;height: 0;max-height: 0;max-width: 0;opacity: 0;overflow: hidden;mso-hide: all; visibility: hidden; width: 0;}.powered-by a {text-decoration: none;}hr {border: 0;border-bottom: 1px solid #f6f6f6;margin: 20px 0;}@media only screen and (max-width: 620px) {table[class=body] h1 {font-size: 28px !important;margin-bottom: 10px !important;} table[class=body] p,table[class=body] ul, table[class=body] ol,table[class=body] td,table[class=body] span,table[class=body] a {font-size: 16px !important;}table[class=body] .wrapper,table[class=body] .article {padding: 10px !important;}table[class=body] .content {padding: 0 !important; }table[class=body] .container {padding: 0 !important;width: 100% !important; }table[class=body] .main {border-left-width: 0 !important;border-radius: 0 !important;border-right-width: 0 !important; }table[class=body] .btn table {width: 100%important;} table[class=body] .btn a {width: 100% !important;}table[class=body] .img-responsive {height: auto !important;max-width: 100% !important;width: auto !important;}}@media all {.ExternalClass { width: 100%; } .ExternalClass, .ExternalClass p,.ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div { line-height: 100%;} .apple-link a { color: inherit !important; font-family: inherit !important; font-size: inherit !important; font-weight: inherit !important; line-height: inherit !important; text-decoration: none !important;} #MessageViewBody a {color: inherit; text-decoration: none; font-size: inherit; font-family: inherit; font-weight: inherit; line-height: inherit; } .btn-primary table td:hover {background-color: #34495e !important; }.btn-primary a:hover {background-color: #34495e !important;border-color: #34495e !important;} }</style></head>";
                string body1 = "<body class=''> <span class='preheader'>Por favor revise este correo.</span> <table role='presentation' border='0' cellpadding='0' cellspacing='0' class='body'><tr><td>&nbsp;</td><td class='container'>";
                string body2 = "<div class='content'><table role='presentation' class='main'><tr><td class='wrapper'><table role='presentation' border='0' cellpadding='0' cellspacing='0'><tr><td><p>" + contenido + "</p><p style='text-align:center;'><img src= " + pathImg1 + "></p><table role='presentation' border='0' cellpadding='0' cellspacing='0' class='btn btn-primary'><tbody><tr><td align='left'><table role='presentation' border='0' cellpadding='0' cellspacing='0'><tbody><tr><td>" + btnLinkSitio + "</td></tr></tbody></table></td></tr></tbody></table>";
                string footer = "<p>" + firma + "</p></td></tr></table></td></tr></table><div class='footer'><table role='presentation' border='0' cellpadding='0' cellspacing='0'><tr><td class='content-block'></td></tr><tr><td class='content-block powered-by'></td></tr></table></div></div></td><td>&nbsp;</td></tr></table></body></html>";

                string plantilla = header + body1 + body2 + footer;
                archivo.WriteLine(plantilla);
                archivo.Close();

                //////-------------------------------------------------------/////
                string pathTxt = string.Format("{0}/{1}", Server.MapPath("~/Content/Asignaciones"), nombrePlantilla + ".txt");
                using (StreamWriter writer = new StreamWriter(pathTxt, false))
                {
                    for (int i = 0; i < values.Count; i++)
                    {
                        writer.WriteLine(values[i].ToString());
                    }
                    writer.Close();
                }
                ///----------------------------------------------------///
                return new JsonResult() { Data = cargarCombo(), JsonRequestBehavior = JsonRequestBehavior.AllowGet, MaxJsonLength = Int32.MaxValue };
            }
            catch (Exception ex)
            {
                string Error;
                Error = ex.Message;

                return Json(new { error = true, message = "Ha ocurrido un error al grabar la plantilla: " + Error }, JsonRequestBehavior.AllowGet);
            }

        }
        public string[] cargarCombo()
        {
            string path1 = string.Format("{0}", Server.MapPath("~/Content/Uploads/Plantillas"));
            string[] ubicacion = Directory.GetFiles(path1);
            string[] archivos = new string[ubicacion.Length];

            for (int i = 0; i < ubicacion.Length; i++)
            {
                archivos[i] = (Path.GetFileName(ubicacion[i]));

            }
            return archivos;
        }

        public ActionResult LeerExcel(string xlsxFile)
        {

            if (xlsxFile != null )
            {
                try
                {
                    xlsxFile = Convert.ToString(TempData["Ruta"]);
                    var excelFile = new FileInfo(xlsxFile);
                    if (!excelFile.Exists)
                        throw new FileNotFoundException($"The file '{excelFile.FullName}' does not exist.", excelFile.FullName);


                    using (var package = new ExcelPackage(excelFile))
                    {
                        ExcelWorksheet workSheet = package.Workbook.Worksheets[0];

                        string[] cols = new string[workSheet.Dimension.End.Column];
                        //string[] cols1 = new string[workSheet.Dimension.End.Column];


                        // Read Excel file column titles (the first worhsheet row)
                        var columnIndexByName = new Dictionary<string, int>();

                        for (int col = 1; col <= workSheet.Dimension.End.Column; col++)
                        {
                            if(workSheet.Cells[1, col].Value != null)
                            {
                                string columnName = workSheet.Cells[1, col].Value.ToString().ToLower();

                                cols[col - 1] = columnName;
                            }

                        }

                        if(cols[0] != null)
                        {
                            ViewData["colsExcel"] = cols;

                            return new JsonResult() { Data = cols, JsonRequestBehavior = JsonRequestBehavior.AllowGet, MaxJsonLength = Int32.MaxValue };
                        }
                        else
                        {
                            cols = new string[] { "" };

                            ViewData["colsExcel"] = cols;

                            return new JsonResult() { Data = cols, JsonRequestBehavior = JsonRequestBehavior.AllowGet, MaxJsonLength = Int32.MaxValue };
                        }


                        //for (int col = 1; col <= workSheet.Dimension.End.Column; col++)
                        //{
                        //    if (workSheet.Cells[1, col].Value != null)
                        //    {
                        //        string columnName = workSheet.Cells[1, col].Value.ToString().ToLower();

                        //        cols1[col - 1] = columnName;
                        //    }
                        //}

                        //if (cols1 != null)
                        //{
                        //    ViewData["colsExcel"] = cols1;
                        //}

                    }
                }
                catch (Exception ex)
                {
                    string Error;
                    Error = ex.Message;

                    return Json(new { success = false, message = "Ha ocurrido un error al leer el archivo excel: " + Error }, JsonRequestBehavior.AllowGet);
                }

            }
            return Json(new { error = true, message = "Error" }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult guardarTag(string tag, List<String> values)
        {
            try
            {
                string nomTag = '[' + tag + ']';
                int aux = 0;
                if (values == null)
                {
                    listaTag.Add(nomTag);
                    aux = 1;
                }
                else
                {
                    if (values.Contains(nomTag))
                    {
                        aux = 1;
                    }
                }

                if (aux != 1)
                {
                    listaTag.Add('[' + tag + ']');
                }
                if (listaTag.Count == 0)
                {
                    listaTag = null;
                }

                return Json(new { success = true, data = listaTag }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                string Error;
                Error = ex.Message;

                return Json(new { error = true, message = "Ha ocurrido un error al guardar el tag: " + Error }, JsonRequestBehavior.AllowGet);
            }

        }


        public ActionResult AsignarTagColumna(string tag, string col)
        {
            try
            {
                string tagcol = tag + "," + col;
                listaTagCol.Add(tagcol);

                return Json(new { success = true, data = listaTagCol }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                string Error;
                Error = ex.Message;

                return Json(new { error = true, message = "Ha ocurrido un error al asignar el tag a la columna: " + Error }, JsonRequestBehavior.AllowGet);
            }


        }

        public JsonResult ConfirmarAsignaciones(List<String> values)
        {
            try
            {
                int numero = values.Count;

                string[] separadas = new string[2 * values.Count];
                string[] nomTag = new string[values.Count];
                string[] nomCol = new string[values.Count];

                for (int i = 0; i < values.Count; i++)
                {
                    if (separadas[0] == null)
                    {
                        separadas = values[i].Split(',').ToArray();
                    }
                    else
                    {
                        separadas = separadas.Concat(values[i].Split(',')).ToArray();
                    }

                }
                int aux = 0;
                int aux1 = 0;
                for (int j = 0; j < separadas.Length; j++)
                {

                    if (j == 0 || j % 2 == 0)
                    {
                        nomTag[aux] = separadas[j];
                        aux++;
                    }
                    else
                    {
                        nomCol[aux1] = separadas[j];
                        aux1++;
                    }
                }

                return Json(new { data = nomTag });
            }
            catch (Exception ex)
            {
                string Error;
                Error = ex.Message;

                return Json(new { error = true, message = "Ha ocurrido un error al confirmar las asignaciones: " + Error }, JsonRequestBehavior.AllowGet);
            }


        }
    }
}