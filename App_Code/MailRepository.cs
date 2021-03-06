﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ActiveUp.Net.Mail;

/// <summary>
/// Summary description for MailRepository
/// </summary>
public class MailRepository
{
    private Imap4Client client;

    public MailRepository(string mailServer, int port, bool ssl, string login, string password)
    {
        if (ssl)
            Client.ConnectSsl(mailServer, port);
        else
            Client.Connect(mailServer, port);
        Client.Login(login, password);
    }

    public IEnumerable<Message> GetAllMails(string mailBox)
    {
        return GetMails(mailBox, "ALL").Cast<Message>();
    }

    public IEnumerable<Message> GetUnreadMails(string mailBox)
    {
        return GetMails(mailBox, "UNSEEN").Cast<Message>();
    }

    protected Imap4Client Client
    {
        get { return client ?? (client = new Imap4Client()); }
    }

    private MessageCollection GetMails(string mailBox, string searchPhrase)
    {
        Mailbox mails = Client.SelectMailbox(mailBox);
        MessageCollection messages = mails.SearchParse(searchPhrase);
        return messages;
    }

    public void MoveTOProcessedFolder(string fromMailBox, string toMailBox, string strMsgId) 
    {
        //Copy and delete from fromMailBox folder to toMailBox folder
        FlagCollection flags;
        Mailbox mails = Client.SelectMailbox(fromMailBox);
        int[] ids = mails.Search("ALL");
        if (ids.Length > 0)
        {
            ActiveUp.Net.Mail.Message msg = null;
            for (var i = 0; i < ids.Length; i++)
            {
                msg = mails.Fetch.MessageObject(ids[i]);
                if( strMsgId.Contains ( Convert.ToString(ids[i]) ) )
                {
                    Client.Command("copy " + ids[i].ToString() + toMailBox );  //copy emails

                    flags = new FlagCollection(); //delete emails
                    flags.Add("Deleted");
                    mails.AddFlags(ids[i], flags);
                }
            }
        }
    }   
}