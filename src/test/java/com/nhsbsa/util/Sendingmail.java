package com.nhsbsa.util;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Properties;

import org.apache.commons.text.StringEscapeUtils;

import jakarta.activation.DataHandler;
import jakarta.activation.DataSource;
import jakarta.activation.FileDataSource;
import jakarta.mail.Address;
import jakarta.mail.Authenticator;
import jakarta.mail.Message;
import jakarta.mail.MessagingException;
import jakarta.mail.Multipart;
import jakarta.mail.PasswordAuthentication;
import jakarta.mail.Session;
import jakarta.mail.Transport;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeBodyPart;
import jakarta.mail.internet.MimeMessage;
import jakarta.mail.internet.MimeMultipart;

import jxl.Sheet;

import jxl.Workbook;

public class Sendingmail {
	public static void sendMail(String html) throws Exception {
		try {
			String[] sendMailTo = null;
			String[] sendMailCc = null;
			String subject = null;
			InternetAddress[] addressTo;
			InternetAddress[] addressCc;
			// Credentials and host
			String host = "smtp.gmail.com";
			final String user = ITestListenerImpl.defaultConfigProperty.get().getProperty("MailFromUser");
			final String team = ITestListenerImpl.defaultConfigProperty.get().getProperty("MailFromTeam");
			final String password = "jotm irmm boyr rxmh";
			// Email message content
			String bodymsg = html;
			// Get the session object
			Properties props = new Properties();
			props.put("mail.smtp.host", "smtp.gmail.com");
			props.put("mail.smtp.port", "587");
			props.put("mail.smtp.auth", "true");
			props.put("mail.smtp.starttls.enable", "true");
			props.put("mail.smtp.ssl.trust", "smtp.gmail.com");

			Session session = Session.getInstance(props, new Authenticator() {
				protected PasswordAuthentication getPasswordAuthentication() {
					return new PasswordAuthentication("avyakthkumarashok@gmail.com", "jotmirmmboyrrxmh");
				}
			});

			// Compose the message
			MimeMessage message = new MimeMessage(session);
			// MailTo
			String mailTo = ITestListenerImpl.defaultConfigProperty.get().getProperty("MailTo");
			String to1 = mailTo;
			if (to1.contains(",")) {
				sendMailTo = to1.split(",");
				addressTo = new InternetAddress[sendMailTo.length];
				for (int i = 0; i < sendMailTo.length; i++) {
					addressTo[i] = new InternetAddress(sendMailTo[i]);
				}
				message.setRecipients(Message.RecipientType.TO, addressTo);
			} else if (!(to1.length() < 0)) {
				InternetAddress[] toAdressArray = InternetAddress.parse(to1);
				message.addRecipients(Message.RecipientType.TO, toAdressArray);
			} else {
				System.out.println("Invalid MailTo");
			}
			// Mailcc
			String mailCc = ITestListenerImpl.defaultConfigProperty.get().getProperty("MailCc");
			String cc = mailCc;
			if (cc.contains(",")) {
				sendMailCc = cc.split(",");
				addressCc = new InternetAddress[sendMailCc.length];
				for (int i = 0; i < sendMailCc.length; i++) {
					addressCc[i] = new InternetAddress(sendMailCc[i]);
				}
				message.setRecipients(Message.RecipientType.CC, addressCc);
			} else if (!(cc.length() < 0)) {
				InternetAddress[] ccAdressArray = InternetAddress.parse(cc);
				message.addRecipients(Message.RecipientType.CC, ccAdressArray);
			} else {
				System.out.println("Invalid MailCc");
			}
			// Print all recepients
			Address[] mssg = message.getAllRecipients();
			for (int i = 0; i < mssg.length; i++) {
				// System.out.println("The recepients are " + mssg[i]);
			}
			// Mailfrom
			message.setFrom(new InternetAddress(user, team));
			// Set Subject
			Calendar cal = Calendar.getInstance();
			DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm");
			String calendarDate = dateFormat.format(cal.getTime());
			subject = ITestListenerImpl.defaultConfigProperty.get().getProperty("AppName")
					+ " [Tech Exercise] Automation Report – " + calendarDate;
			message.setSubject(subject);
			String cleaned = bodymsg.replace("&nbsp", " ");
			message.setContent(cleaned, "text/html; charset=UTF-8");

			// send the message
			final long startTime = System.currentTimeMillis();
			Transport.send(message, message.getAllRecipients());
			long endTime = System.currentTimeMillis();
			long totalTime = endTime - startTime;
			System.out.println("==========>Mail Sent Successfully in " + (totalTime / 1000) % 60 + " secs <==========");
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("Exception in SendMail " + e.getMessage());
		}
	}
}
