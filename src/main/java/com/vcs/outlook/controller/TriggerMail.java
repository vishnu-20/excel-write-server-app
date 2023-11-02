package com.vcs.outlook.controller;

import java.awt.Desktop;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URLEncoder;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class TriggerMail {

	@GetMapping("/openMail/{email}/{bookingNumber}")
	void openOutlook(@PathVariable String email, @PathVariable String bookingNumber) {
		try {
			String subject = "Hello, Outlook!";
			String body = "This is a test email sent from a Java program using Outlook.";
			subject = URLEncoder.encode(subject, "UTF-8");
			body = URLEncoder.encode(body, "UTF-8");

			URI mailtoURI = new URI("mailto:" + email + "?subject=" + subject + "&body=" + body);

			if (Desktop.isDesktopSupported()) {
				Desktop desktop = Desktop.getDesktop();

				if (desktop.isSupported(Desktop.Action.MAIL)) {
					desktop.mail(mailtoURI);
					System.out.println("Opening Outlook to compose an email...");
				}
			} else {
				System.out.println("Desktop is not supported on this platform.");
			}
		} catch (IOException | URISyntaxException e) {
			e.printStackTrace();
		}

	}

}
