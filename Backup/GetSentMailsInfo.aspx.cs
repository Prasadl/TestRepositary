using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data.SqlClient;
using System.Configuration;
using ActiveUp.Net.Mail;
using System.Data;
using ExcelLibrary.SpreadSheet;
using QiHe.CodeLib;
using ExcelLibrary.CompoundDocumentFormat;
using System.IO;


public partial class GetSentMailsInfo : System.Web.UI.Page
{
   static string strUsrName, strPwd;

    protected void Page_Load(object sender, EventArgs e)
    { 
        divMsg.InnerHtml = "";
        if(!IsPostBack)
        {
            txtUsrName.Text = "";
            txtpwd.Text = "";
            tblLogin.Visible = true;
            tblEmails.Visible = false;
            strUsrName = "";
            strPwd = "";
        }
        if (strUsrName == "" || strPwd == "")
        {
            tblLogin.Visible = true;
            tblEmails.Visible = false;
        }
        else
        {
            tblLogin.Visible = false;
            tblEmails.Visible = true;
        }
    }

    protected void btnScanEmails_onClick(object sender, EventArgs e)
    {
        try
        {
            ReadSentEmails();
        }
        catch (Exception ex)
        {
            //btnScanEmails.Visible = false;
            //btnExportToExcel.Visible = false;
            tblLogin.Visible = false;
            tblEmails.Visible = false;
            divError.InnerText = Convert.ToString(ex);
        }
    }

    protected void ReadSentEmails()
    {
        string sentDate = "", companyWebsite = "", companyName = "", clientEmailId = "", clientName = "", fromEmailId = "";  //companyPhone, clientPhone
        string strEmail = "", strEmailBody = "", strEmailids = "", strWebsite = "";

        try
        {

            var mailRepository = new MailRepository(
                                    ConfigurationManager.ConnectionStrings["ServerName"].ConnectionString,
                                    int.Parse(ConfigurationManager.ConnectionStrings["PortNo"].ConnectionString),
                                    true,
                                    strUsrName,
                                    strPwd
                                );
            var emailList = mailRepository.GetAllMails("Sent");

            foreach (Message email in emailList)
            {
                foreach (Address emailid in email.To)
                {
                    if (strEmailids == "")
                        strEmailids += emailid;
                    if (strWebsite == "")
                        strWebsite = Convert.ToString(emailid).Substring((Convert.ToString(emailid)).IndexOf("@") + 1);
                }

                companyWebsite += "www." + strWebsite + "|";
                strWebsite = "";

                //foreach (Address emailid in email.Cc)
                //{  }
                //foreach (Address emailid in email.Bcc)
                //{  }

                clientEmailId += strEmailids + "|"; //Contains all To, CC, BCC email ids
                strEmailids = "";

                sentDate += email.Date + "|";

                fromEmailId += Convert.ToString(email.From) + "|";

                //strEmail += "<br><b>ReceivedDate :</b> :" + email.ReceivedDate;
                //strEmail += "<br><b>MessageId</b> :" + email.MessageId;
                //strEmail += "<br><b>Subject</b> :" + email.Subject;
                strEmailBody = Convert.ToString(email.BodyText.Text);
                if (strEmailBody.Contains("Good Morning"))
                {
                    int i = strEmailBody.IndexOf("Good Morning");
                    string strTempEBody = strEmailBody.Substring(i);
                    int j = strTempEBody.IndexOf(",");
                    int k = ("Good Morning ").Length;
                    strEmail += "<br><b>Client Name</b> :" + strEmailBody.Substring((i + k), (j - k));

                    clientName += strEmailBody.Substring((i + k), (j - k)) + "|";
                }
                else
                    clientName += "|";

                if (strEmailBody.Contains("partnering with companies such as"))
                {
                    int i = strEmailBody.IndexOf("partnering with companies such as");
                    string strTempEBody = strEmailBody.Substring(i);
                    int j = strTempEBody.IndexOf(",");
                    int k = ("partnering with companies such as ").Length;
                    strEmail += "<br><b>Company Name</b> :" + strEmailBody.Substring((i + k), (j - k));

                    companyName += strEmailBody.Substring((i + k), (j - k)) + "|";
                }
                else
                    companyName += "|";

                if (email.Attachments.Count > 0)
                {
                    foreach (MimePart attachment in email.Attachments)
                    {
                        strEmail += "<p>Attachment:" + attachment.ContentName + "  " + attachment.ContentType.MimeType + "</p><br>";
                    }
                }
            }
            insertRecords(sentDate, companyWebsite, companyName, clientEmailId, clientName, fromEmailId);

            divMsg.InnerHtml = "<b> <font color=\"red\"> All Emails are scanned successfully. </font> </b>";
        }
        catch (Exception ex)
        {
            //btnScanEmails.Visible = false;
            //btnExportToExcel.Visible = false;
            tblLogin.Visible = false;
            tblEmails.Visible = false;
            if (ex.ToString().Contains("failed"))
                divError.InnerHtml = "<b> <font color=\"red\"> Please enter valid Email Id and Password. </font> </b>";  // +ex.ToString();         
            else
                divError.InnerText = Convert.ToString(ex);
        }
    }

    protected void insertRecords(string sentDate, string companyWebsite, string companyName, string clientEmailId, string clientName, string fromEmailId)
    {
        try
        {
            using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["SQLServer"].ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand("sp_insertEmailsInfo", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("@sentDate", SqlDbType.VarChar).Value = sentDate;
                    cmd.Parameters.Add("@companyWebsite", SqlDbType.VarChar).Value = companyWebsite;
                    cmd.Parameters.Add("@companyName", SqlDbType.VarChar).Value = companyName;
                    cmd.Parameters.Add("@clientEmailId", SqlDbType.VarChar).Value = clientEmailId;
                    cmd.Parameters.Add("@clientName", SqlDbType.VarChar).Value = clientName;
                    cmd.Parameters.Add("@fromEmailId", SqlDbType.VarChar).Value = fromEmailId;

                    con.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }
        catch (Exception ex)
        {
            //btnScanEmails.Visible = false;
            //btnExportToExcel.Visible = false;
            tblLogin.Visible = false;
            tblEmails.Visible = false;
            divError.InnerText = Convert.ToString(ex);
        }
    }

    

    protected void exportToExcel()
    {
        string file = "";
        try
        {
            DateTime dtFrom = DateTime.Today, dtTo = DateTime.Today;

            if (Convert.ToString(txtFromDate.Text) != "" && Convert.ToString(txtFromDate.Text) != null)
                dtFrom = DateTime.Parse(txtFromDate.Text);
            if (Convert.ToString(txtToDate.Text) != "" && Convert.ToString(txtToDate.Text) != null)
             dtTo = DateTime.Parse(txtToDate.Text);

            if ( dtTo < dtFrom )
                divError.InnerHtml = "<b> <font color=\"red\"> Please select proper To and From Dates. </font> </b>";
            
            System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
            file = ConfigurationManager.ConnectionStrings["FilePath"].ConnectionString + "Report.xls";  //"C:\\Report.xls";  // +DateTime.Now.ToString().Replace("/", "_").Replace(":", "_") + ".xls";
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("First Sheet");

            string[] colNames = { "Id", "Sent Date", "Company Website", "Company Name", "Client EmailId", "Client Name", "Client Phone" };  // "Company Phone"

            for (int i = 0; i < colNames.Length; i++)
            {
                worksheet.Cells[0, i] = new Cell(colNames[i]);
                worksheet.Cells.ColumnWidth[0, (ushort)i] = 5000;
            }

            DataSet ds = new DataSet("dsEmailsInfo");
            using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["SQLServer"].ConnectionString))
            {
                SqlCommand cmd = new SqlCommand("sp_getEmailsInfo", con);
                if (dtFrom == DateTime.Today && dtTo == DateTime.Today) //From and To dates are null for getting all records
                {
                    cmd.Parameters.AddWithValue("@startDate", DBNull.Value);
                    cmd.Parameters.AddWithValue("@endDate", DBNull.Value);
                }
                else if (dtFrom == DateTime.Today && dtTo != DateTime.Today) //From and To dates are null for getting all records
                {
                    cmd.Parameters.AddWithValue("@startDate", DBNull.Value);
                    cmd.Parameters.AddWithValue("@endDate", dtTo);
                }
                else
                {
                    cmd.Parameters.AddWithValue("@startDate", dtFrom);
                    cmd.Parameters.AddWithValue("@endDate", dtTo);
                }
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(ds, "EmailsInfo");

                DataTable dt = ds.Tables[0];
                int intRowNo = 1;
	   int intSrNo = 1;
                foreach (DataRow row in dt.Rows)
                {
                    intRowNo++;
                    int i = 0;
                    foreach (DataColumn col in dt.Columns)
                    {
	           if ( i == 0 )
	           {
		worksheet.Cells[intRowNo, i++] = new Cell(Convert.ToString( intSrNo++ ));
                          worksheet.Cells.ColumnWidth[(ushort)intRowNo, (ushort)i] = 1000;
	           }
	           else
	           {
		worksheet.Cells[intRowNo, i++] = new Cell(Convert.ToString(row[col]));
                         worksheet.Cells.ColumnWidth[(ushort)intRowNo, (ushort)i] = 6000;
	           }
                    }
                }
            }           
            workbook.Worksheets.Add(worksheet);
            //-----------------------------------------
            worksheet = new Worksheet("Second Sheet");
            for (int i = 0; i < 550; i++)
            {
                worksheet.Cells[i, 0] = new Cell(i);
            }
            workbook.Worksheets.Add(worksheet);
            //-----------------------------------------------------------
            workbook.Save(file);

            string name = file.Substring((file.LastIndexOf("\\") + 1));
            string fileName = "Report" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_") + ".xls";
            string type = "application/vnd.ms-excel";
            if (true)
            {
                Response.AppendHeader("content-disposition", "attachment; filename=" + fileName);
            }
            if (type != "")
            {
                Response.ContentType = type;
                Response.WriteFile(file);
                Response.End();
            }
        }
        catch (Exception ex)
        {
            //btnScanEmails.Visible = false;
            //btnExportToExcel.Visible = false;
            tblLogin.Visible = false;
            tblEmails.Visible = false;
            divError.InnerText = Convert.ToString(ex);
        }
    }    
    
    protected void btnExportToExcel_onClick(object sender, EventArgs e)
    {
        try 
	    {
            exportToExcel();
	    }
	    catch (Exception ex)
	    {
            //btnScanEmails.Visible = false;
            //btnExportToExcel.Visible = false;
            tblLogin.Visible = false;
            tblEmails.Visible = false;
            divError.InnerText = Convert.ToString(ex);
	    }
    }

    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        try 
	    {
            tblLogin.Visible = false;
            tblEmails.Visible = true;
            strUsrName = Convert.ToString(txtUsrName.Text);
            strPwd  = Convert.ToString(txtpwd.Text);

            if (strUsrName != ConfigurationManager.ConnectionStrings["EmailId"].ConnectionString)
            {
                strUsrName = "";
                strPwd = "";
                throw new Exception();
            }
            else
            {
            }

	    }
	    catch (Exception ex)
	    {
            //btnScanEmails.Visible = false;
            //btnExportToExcel.Visible = false;
            //divError.InnerText = Convert.ToString(ex);
            tblLogin.Visible = false;
            tblEmails.Visible = false;
              divError.InnerHtml = "<b> <font color=\"red\"> Please enter valid Email Id and Password. </font> </b>";  
            	
	    }

    }
}

