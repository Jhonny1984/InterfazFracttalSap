using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using RestSharp;
using Newtonsoft.Json;
using System.Data.SqlClient;
using System.Data;
using System.ComponentModel;
using HawkNet;
using System.Web.UI;
using System.Web;
using System.Security.Permissions;
using MSScriptControl;
using System.Threading;
using System.Net;
using System.Diagnostics;
using System.IO;
using System.Web;
using System.Net.Mail;
using Aspose.Email.Clients.ActiveSync.TransportLayer;
using MailChimp.Net.Models;

namespace ProInterfaceFracttal
{
    public class EnviaCorreo
    {
        public static void EnviaCorreos(string err, string NumAtCard, string DocNum) {



            String FROM = "fracttal.transflesa@macizo.com.gt";
            String FROMNAME = "ProInterface Fracttal";

            String TO = "geovani.ajozal@macizo.com.gt";
            String TO2 = "fabian.canon@macizo.com.gt";
            String CC = "ealvarez@macizo.com.gt";
            String CC2 = "jhonny.rueda@macizo.com.gt";

            String SMTP_USERNAME = "fracttal.transflesa@macizo.com.gt";

            String SMTP_PASSWORD = "Macizo00*";

            String CONFIGSET = "ConfigSet";

            String HOST = "mail.macizo.com.gt";

            int PORT = 26;


            String SUBJECT =
                "Oferta de Venta ProInterfaz Fracttal";


            String BODY =

                BODY = @"
                
                <h1>ProInterface Fracttal</h1>
                <p>Realizando Ofertas de Venta
                    <a href='https://macizo.com.gt'>Macizo</a> proInterface
                </p>";

            // SW.WriteLine(mensajeB, mensajeA, mensaje);
            ;

            // Console.WriteLine ("el cuerpo del texto es: " + mensajeB );


            // SW.Close();

            MailMessage message = new MailMessage();
            message.IsBodyHtml = true;
            message.From = new MailAddress(FROM, FROMNAME);
            message.To.Add(new MailAddress(TO));
            message.To.Add(new MailAddress(TO2));
            message.CC.Add(new MailAddress(CC));
            message.CC.Add(new MailAddress(CC2));
            message.Subject = SUBJECT;
            message.Body = BODY + err + " --- En la orden de trabajo numero: " + NumAtCard + " - " + DocNum;


            message.Headers.Add("X-SES-CONFIGURATION-SET", CONFIGSET);

            using (var client = new System.Net.Mail.SmtpClient(HOST, PORT))
            {

                client.Credentials =
                    new NetworkCredential(SMTP_USERNAME, SMTP_PASSWORD);

                client.EnableSsl = true;


                try
                {
                  //  Console.WriteLine("Intentando enviar el correo");
                    client.Send(message);
                    Console.WriteLine("Correo ha sido enviado correctamente!");
                    //Console.Read();
                    //Environment.Exit(0);
                }
                catch (Exception ex1)
                {
                    Console.WriteLine("El correo no ha sido enviado.");
                    Console.WriteLine("Error del mensaje: " + ex1.Message);
                    //Console.Read();
                    //Environment.Exit(0);
                }

            }




        }


    }
}
