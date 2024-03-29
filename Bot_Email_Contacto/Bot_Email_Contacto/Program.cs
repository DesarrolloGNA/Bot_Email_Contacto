﻿
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using AE.Net.Mail;

namespace Bot_Email_Contacto
{
    class Program
    {
        static void Main(string[] args)
        {

            do
            {
                try
                {

                    Console.Clear();
                    AuthMethods method = AuthMethods.Login;
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    DateTime Dia_Actual = DateTime.Now;

                    Console.WriteLine(" Busqueda de Correos - " + Dia_Actual.ToString("yyyy-MM-dd hh:mm:ss"));
                
                    using (var imap = new ImapClient("outlook.office365.com", "botsgna@outlook.com", "w2003ejf103", method, 993, true))
                    {
                        SearchCondition condition = new SearchCondition();
                        condition.Value = SearchCondition.To("contacto@gna.cl");

                        Lazy<MailMessage>[] msgs = imap.SearchMessages(condition);
             
                        int Cantidad = msgs.Length;

                        for (int i = 1; i < Cantidad; i++)
                        {


                            Console.WriteLine((i.ToString() + " de " + Cantidad.ToString() + "---" + msgs[i].Value.From.ToString()));

                            if (msgs[i].Value.Date.ToShortDateString() == Dia_Actual.ToShortDateString())
                            {

                                if ((!msgs[i].Value.From.ToString().Contains("gna.cl") && !msgs[i].Value.From.ToString().Contains("garcianadal.cl") && !msgs[i].Value.From.ToString().Contains("noreply")) || ((!msgs[i].Value.From.ToString().Contains("gna.cl") || !msgs[i].Value.From.ToString().Contains("garcianadal.cl")) && msgs[i].Value.Subject.ToUpper().Contains("CELULAR ESTEBAN")))
                                {

                                    try
                                    {

                               
                                    Console.WriteLine((msgs[i].Value.From.ToString() + " - " + msgs[i].Value.Date.ToString("yyyy-MM-dd hh:mm:ss")));

                                    string Remitente = msgs[i].Value.From.ToString();
                                    DateTime Fecha_Carga = msgs[i].Value.Date;
                                    string Email_Asunto = msgs[i].Value.Subject;
                                    string Email_Mensaje = msgs[i].Value.Body.ToString();

                                    Grabar_Email(Fecha_Carga, Remitente, Email_Asunto, Email_Mensaje);

                                    }
                                    catch (Exception EX)
                                    {
                                        Console.WriteLine(EX.Message.ToString());
                                    }


                                }

                            }

                        }
                    }
                    Dia_Actual = DateTime.Now;
                    Console.WriteLine("Finalizo  la busqueda, Reintentara en 15 minutos mas mas - ("+ Dia_Actual.AddMinutes(15).ToString("yyyy-MM-dd hh:mm:ss") + ")");
                    Thread.Sleep(900000);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message.ToString());
                }

            } while (true);

        }

        public static void Grabar_Email(DateTime Fecha_Carga,string Email_Remitente,string Email_Asunto,string Email_Mensaje)
        {

            string connstring = @"Data Source=192.168.0.5; Initial Catalog=GESTIONES_PREJUDICIALES; Persist Security Info=True; User ID=sa; Password=w2003ejf103;";
  
            using (SqlConnection conexion = new SqlConnection(connstring))
            {

                SqlCommand cmd = new SqlCommand("SP_INBOUND_EMAIL_CREATE", conexion);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@Fecha_Carga", Fecha_Carga));
                cmd.Parameters.Add(new SqlParameter("@Email_Remitente", Email_Remitente));
                cmd.Parameters.Add(new SqlParameter("@Email_Asunto", Email_Asunto));
                cmd.Parameters.Add(new SqlParameter("@Email_Mensaje", Email_Mensaje));

                conexion.Open();
                cmd.ExecuteNonQuery();
                conexion.Close();

            }


        }
    }
}
