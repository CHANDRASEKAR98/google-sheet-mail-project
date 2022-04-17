# Google Sheet Mail Project
In this project, based on the google sheet data, the mail is sent to the recipients. For a real time case, let's take an example of sending birthday wishes to the employees in an organization. When the program is executed, birthday wishes are sent to the employees who are celebrating their birthday on this day.

## Prerequisite
### Softwares / Tools required
- JDK 8
- Java IDE (Ex: Eclipse, VsCode)
- A Google Account
### Java dependencies / JARs requried
- [Google API Client](https://jar-download.com/artifact-search/google-api-client)
- [Google API Service](https://jar-download.com/artifacts/com.google.apis/google-api-services-drive)
- [Google API Sheets](https://jar-download.com/artifact-search/google-api-services-sheets)
- [Google Oauth Client Java6](https://jar-download.com/artifacts/com.google.oauth-client/google-oauth-client-java6)
- [Javax Mail API](https://jar-download.com/artifacts/com.sun.mail/javax.mail)

## Road Map
1. Creating a Google Account
2. Enabling a Google API Service
3. Downloading Client Secret JSON file
4. Creating a Google Sheet
5. Accessing Google Sheet data using Java
6. Using Javax mail to send emails
7. Running the Java Program and verifying the output

## Installation
### JDK 8 Installation
Download and install Java 8 JDK on your system if not already done.
You can download Java 8 from the official [Oracle Website](https://www.oracle.com/in/java/technologies/javase/javase8-archive-downloads.html) 

Once downloaded, Install and setup your environemntal variables for Java 8 properly. If you don't know how to setup your environmental variables, kindly follow the steps given in this [link](https://www.javatpoint.com/how-to-set-path-in-java).

To check the version of Java installed in your system, kindly execute the following command on your CMD.

```bash
  java -version
```
The output will be displayed like the following.

```bash
  java version "1.8.0_101"
  Java(TM) SE Runtime Environment (build 1.8.0_101-b13)
  Java HotSpot(TM) 64-Bit Server VM (build 25.101-b13, mixed mode)
```

## Accessing Google Sheet using Java Program
To acesss the google sheet data using Java program, you have to complete the steps given in this [Link](https://github.com/CHANDRASEKAR98/google-sheet-api-project/edit/main/README.md#google-sheet-api-project)

Once you have completed those steps, you will have access to the google sheet data within your Java Program.

## Using Javax Mail to send Emails
To send mails to the recipients, you'll have to walk through the steps given in this [Link](https://github.com/CHANDRASEKAR98/java-email-project#java-email-project)

Now you are eligible to do the following
  - Accessing Google Sheet Data in Java
  - Sending Mail in Java

## Integrating the Code
Let us integrate the programs you have completed in the above two steps inorder to send Emails to the recipients based on the google sheet Data.

For example, If a person is celebrating his birthday on this day. By integrating the above two Java programs, you can send an email to the person celebrating his birthday.
The following logic is framed to fetch the google sheet data and the value is converted to the type List.

```bash
  ValueRange response = service.spreadsheets().values()
              .get(spreadsheetId, range)
              .execute();
  List<List<Object>> values = response.getValues();
```

Now by streaming the list `values`, you can send emails by a condition check.

```bash
  values.stream.forEach(row -> {
    // condition is checked here and mail is sent accordingly
  });
```

## Run the Java Program
Finally after the code integration, you are all set to run the Java Program.

Let's say if someone in your organization has birthday today (Let's assume that today's date is 2022/04/17). On running the Java program, a condition is checked for the date given the google sheet and today's date. If the condition matches, a birthday mail will be sent.

Refer to the below image for the Google Sheet data.

![]()

Now Let's run the Java Program and see if the mail is sent to the person1.

Yeah!! you did it. The mail has been successfully sent to the person1's mail address. Hope he'll be very happy on receiving the birthday email on his inbox.

Refer the below image for the mail output.

![]()

## Acknowledgement
- [Sending Email in Java](https://www.baeldung.com/java-email)
- [Convert List to Array - Java 8](https://www.geeksforgeeks.org/convert-list-to-array-in-java/)
- [Google Developer Console](https://console.cloud.google.com/)
- [Java Quickstart Sheet API](https://developers.google.com/sheets/api/quickstart/java)
- [JAR Downloads](https://jar-download.com/)

## Authors
- [@Chandrasekar Balakumar](https://github.com/CHANDRASEKAR98)

## Feedback
If you have any feedback, please reach out to me [@Chandrasekar Balakumar](https://www.linkedin.com/in/chandrasekarbalakumar98/) on LinkedIn.
