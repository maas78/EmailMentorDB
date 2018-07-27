using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.IO;
using ADODB;
namespace EmailMentorDB
{
  public partial class frmMain : Form
  {
   string[] args=null;
    public frmMain(string[] startup_args)
    {
      InitializeComponent();
    //  if (startup_args.Length != 1)
     //   MessageBox.Show ("Vous devez spécifier le chemin de la base de donnée!") ;
    }

    private void frmMain_Load(object sender, EventArgs e)
    {
      ADODB.Recordset rsMail = null;
      ADODB.Recordset rsContract = null;
      ADODB.Recordset rsMailSettings = null;
      ADODB.Connection connDb = null;
      bool solde=false;
      string destinataire="";
      bool contratEchu=false;
      string CONN_DB ="";
   //   if (args==null)
    //      Application.Exit();
      if (args == null)
          CONN_DB = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\lementor\StudiosUnis\LeMentorÉlèveTables.mdb;";
      else
          CONN_DB = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + args;


      try
      {

        connDb = new ADODB.Connection();
        rsMail = new ADODB.Recordset();
        rsMailSettings = new ADODB.Recordset();
        rsContract = new ADODB.Recordset();
        System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
        client.UseDefaultCredentials = false;
        client.EnableSsl = true;
        client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;

        System.Net.Mail.MailMessage msg = null;
        connDb.Open(CONN_DB);
        rsMail.Open("select * from emailmessage", connDb, CursorTypeEnum.adOpenKeyset, LockTypeEnum.adLockOptimistic);
        rsMailSettings.Open("select * from cie", connDb, CursorTypeEnum.adOpenKeyset, LockTypeEnum.adLockOptimistic);
        if (!rsMail.EOF)
          rsMail.MoveFirst();
        if (!rsMailSettings.EOF)
        {
          rsMailSettings.MoveFirst();
          client.Port = rsMailSettings.Fields["PortSmtp"].Value; // rsMailSettings.Fields["PortSmtp"].Value;
          client.Host = rsMailSettings.Fields["ServerSmtp"].Value.ToString();
          client.Credentials = new System.Net.NetworkCredential(rsMailSettings.Fields["E-mail"].Value.ToString(), rsMailSettings.Fields["MotdePasseE-mail"].Value);

        }

        while (!rsMail.EOF)
        {
          solde = rsMail.Fields["solde"].Value;
          contratEchu = rsMail.Fields["contratechu"].Value;
          
          //destinataire = rsMail.Fields["destinataire"].Value.ToString();

         // destinataire = "dquirion78@@gmail.com";
          msg = new System.Net.Mail.MailMessage();   
          msg.From = new MailAddress(rsMailSettings.Fields["e-mail"].Value.ToString());
          msg.Subject = rsMail.Fields["sujet"].Value.ToString();
          msg.Body = rsMail.Fields["texte"].Value.ToString();
          msg.IsBodyHtml = false;

          if (string.IsNullOrEmpty(rsMail.Fields["destinataire"].Value.ToString()))
            throw new Exception("Destinataire Manquant!");
          else
          {
              msg.To.Add(destinataire);

          }
          if (!string.IsNullOrEmpty(rsMail.Fields["cc"].Value.ToString()))
          {
            msg.CC.Add(rsMail.Fields["cc"].Value.ToString());
          }

          if (!string.IsNullOrEmpty(rsMail.Fields["cci"].Value.ToString()))
          {

            msg.Bcc.Add(rsMail.Fields["cci"].Value.ToString());
          }
          if (rsMail.Fields["fichier"].Value.ToString().Length > 3)
          {
            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(rsMail.Fields["fichier"].Value.ToString());
            msg.Attachments.Add(attachment);

          }
          client.Send(msg);
         // if (rsMail.Fields["fichier"].Value.ToString().Length > 3)
         //   File.Delete(rsMail.Fields["fichier"].Value.ToString());
          msg = null;
          rsContract.Open("select * from GA_Eleve_Cours where [noidcours] =" + rsMail.Fields["noidcours"].Value.ToString(), connDb, CursorTypeEnum.adOpenKeyset, LockTypeEnum.adLockOptimistic);
          if (!rsContract.EOF)
          {
              rsContract.MoveFirst();
              if (rsMail.Fields["ContratEchu"].Value == true)
                  rsContract.Fields["NotficationContratEchu"].Value = true;
              if (rsMail.Fields["solde"].Value == true)
              {
                  rsContract.Fields["DernDateNotificationsolde"].Value = rsMail.Fields["DernDateNotificationsolde"].Value;
                  rsContract.Fields["NbreNotificationsolde"].Value = rsMail.Fields["NbreNotificationsolde"].Value;
              }
              rsContract.Update();
          }
          rsContract.Close();
          rsMail.Delete();
          rsMail.MoveNext();
          if (!rsMail.EOF)
            System.Threading.Thread.Sleep(5000);//limite les chances d'être taggé comme spam par le fournisseur d'envoi en envoyant pas trop vite les emails.
          //msg.CC.Add(Email);
        }


        /*
                if (Attachments.Count() > 0)
                {
                  foreach (var item in Attachments)
                  {
                    if (!string.IsNullOrEmpty(item))
                    {
                      System.Net.Mail.Attachment attachment;
                      attachment = new System.Net.Mail.Attachment(item);
                      msg.Attachments.Add(attachment);
                    }
                  }
                }
                */
        if (((solde == false && contratEchu == false) || (rsMail.RecordCount == 1) ) && destinataire.Length>0)
        {
          MessageBox.Show("Message Envoyé : à " + destinataire);
          Application.Exit();
        }

        else
        {
        //  MessageBox.Show("Message Envoyé : à " + rsMail.RecordCount.ToString() + " clients.");
        }
      }
      catch (Exception ex)
      {
       
        MessageBox.Show(ex.Message);
        MessageBox.Show("ÉCHEC d'envoi de courier !");
        Application.Exit();
      }

      try
      {
          if (rsMail.State != 0)
        {
            rsMail.Close();  
        }
       
        rsMail = null;
        if (rsMailSettings.State != 0)
        {
            rsMailSettings.Close();
        }
        rsMailSettings = null;
        if (rsContract.State!=0)
        {
            rsContract.Close();
        }
        rsContract = null;
        Application.Exit();
      }
      catch
      {
          Application.Exit();
      }
    }
  }
}
