package Horizon.DailyReport;

import java.io.InputStream;
import java.io.IOException;
import java.io.FileInputStream;
import java.util.Properties;
import javax.mail.MessagingException;
import javax.mail.Transport;
import javax.mail.Message;
import javax.mail.Address;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.mail.Session;
import org.apache.log4j.PropertyConfigurator;
import org.apache.log4j.Logger;

public class SendEmail
{
    static Logger log;
    
    static {
        SendEmail.log = Logger.getLogger(SendEmail.class.getName());
    }
    
    public static void sendMail(final String[] args) {
        final Properties prop = getProperties();
        PropertyConfigurator.configure(prop.getProperty("log4jConfPath"));
        final String from = prop.getProperty("FROM");
        final Properties properties = System.getProperties();
        properties.setProperty("mail.smtp.host", prop.getProperty("SMTP_HOST"));
        final Session session = Session.getDefaultInstance(properties);
        try {
            final MimeMessage message = new MimeMessage(session);
            message.setFrom((Address)new InternetAddress(from));
            final String[] receipentList = getToList();
            final int count = receipentList.length;
            final Address[] addresses = new Address[count];
            for (int index = 0; index < count; ++index) {
                addresses[index] = (Address)new InternetAddress(receipentList[index]);
            }
            for (int index = 0; index < addresses.length; ++index) {}
            message.addRecipients(Message.RecipientType.TO, addresses);
            message.setSubject("Daily Deployments/Issue/Task status on " + args[0]);
            final String mailBody = "<div><font size=\"2\" color=\"black\" face=\"Calibri\">Hi All,<br><br>Below is the status of daily Deployments/Issue/Task on <font color=black>" + args[0] + ". </font><br><br>" + "</div>" + "<table   bgcolor=\"#e6e6ff\" cellspacing=\"-1\" style=\"border: 1px solid black;\">" + "<tr>" + "<td >" + "<table cellspacing=\"-1\" ><tr><td valign=\"top\" ><font size=\"2\" color=\"black\" face=\"Calibri\">" + args[1] + "</td></tr></table>" + "</td>" + "<td >" + "<table cellspacing=\"-1\">" + "<tr>" + "<td valign=\"top\" style=\"border: 1px solid black; padding-right: 10px; padding-left: 5px;\"><font size=\"2\" color=\"black\" face=\"Calibri\"> Dev Deployments </td>" + "<td valign=\"top\" style=\"border: 1px solid black; padding-right: 10px; padding-left: 5px;\"><font size=\"2\" color=\"black\" face=\"Calibri\"> Prod/QA Deployment </td>" + "<!--<td valign=\"top\" style=\"border: 1px solid black; padding-right: 10px; padding-left: 5px;\"><font size=\"2\" color=\"black\" face=\"Calibri\"> Dev Daily regular tasks </td>-->" + "<td valign=\"top\" style=\"border: 1px solid black; padding-right: 10px; padding-left: 5px;\"><font size=\"2\" color=\"black\" face=\"Calibri\"> Other issue /Support </td>" + "</tr>" + "<tr bgcolor=\"#ffefcc\">" + "<td valign=\"top\" style=\"border: 1px solid black; padding-right: 10px; padding-left: 5px;\">" + args[2] + "</td>" + "<td valign=\"top\" style=\"border: 1px solid black; padding-right: 10px; padding-left: 5px;\">" + args[3] + "</td>" + "<!--<td valign=\"top\" style=\"border: 1px solid black; padding-right: 10px; padding-left: 5px;\">-->" + "<td valign=\"top\" style=\"border: 1px solid black; padding-right: 10px; padding-left: 5px;\">" + args[5] + "</td>" + "</tr>" + "</table>" + "</td>" + "</tr>" + "</table>" + "<br>Thanks,<br>" + "HORIZON Deployment Team";
            message.setContent((Object)mailBody, "text/html");
            Transport.send((Message)message);
            SendEmail.log.info((Object)"Sent message successfully....");
        }
        catch (MessagingException mex) {
            SendEmail.log.error((Object)"", (Throwable)mex);
        }
    }
    
    public static String[] getToList() {
        final Properties prop = getProperties();
        final String toList = prop.getProperty("TO_LIST");
        final String[] receipentList = toList.split(",");
        return receipentList;
    }
    
    public static Properties getProperties() {
        final Properties prop = new Properties();
        InputStream input = null;
        try {
        	input = new FileInputStream("D:\\Test\\config\\config.properties");
           // input = new FileInputStream("./config/config.properties");
            prop.load(input);
        }
        catch (IOException e) {
            e.printStackTrace();
        }
        return prop;
    }
}
