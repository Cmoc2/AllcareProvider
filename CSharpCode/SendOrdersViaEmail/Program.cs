internal class Program
{
    private static void Main(string[] args)
    {
        //input directory path
        string directoryPath = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Documents\Automations\Orders\ToSend");
        Console.WriteLine($"Monitoring directory: {directoryPath}");
        //set output directory path
        string outputDirectoryPath = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Documents\Automations\Orders\Sent");
        Console.WriteLine($"Output directory: {outputDirectoryPath}");

        //if directories do not exist, create it:
        //input Directory
        if (!Directory.Exists(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        //output Directory
        if (!Directory.Exists(outputDirectoryPath))
        {
            Directory.CreateDirectory(outputDirectoryPath);
        }

        //A dictionary of email mappings, mapping file names or patterns to email addresses.
        Dictionary<string, string> emailMappings = new Dictionary<string, string>
        {
            { "Order - BP", "baldwin.park@allcareprovider.com" },
            { "Order - LA", "la.metro@allcareprovider.com" },
            { "Order - WLA", "la.metro@allcareprovider.com" },
            { "Order - D", "downey@allcareprovider.com" },
            { "Order - SB", "south.bay@allcareprovider.com" },
            { "Order - V", "valley@allcareprovider.com" },
            { "Order - F", "fontana@allcareprovider.com" },
            { "Order - OC", "orange.county@allcareprovider.com" },
            { "Order - CL", "carelon@allcareprovider.com" },
            { "Order - Test", "christian.ortiz@allcareprovider.com" }
        };

        //emailing credentials
        string smtpServer = "smtp.office365.com";
        int smtpPort = 587;

        //Load previously saved SMTP credentials if they exist.
        string credentialsPath = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Documents\Automations\Orders\smtp_credentials.txt");
        string? smtpUser = null;
        string? smtpPassword = null;

        if (File.Exists(credentialsPath))
        {
            var lines = File.ReadAllLines(credentialsPath);
            if (lines.Length == 2)
            {
                string encryptedSmtpUser = lines[0].Split(':')[1];
                string encryptedSmtpPassword = lines[1].Split(':')[1];
                smtpUser = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(encryptedSmtpUser));
                smtpPassword = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(encryptedSmtpPassword));
                if (string.IsNullOrWhiteSpace(smtpUser) || string.IsNullOrWhiteSpace(smtpPassword))
                {
                    Console.WriteLine("SMTP credentials are required.");
                    Environment.Exit(1);
                }
                Console.WriteLine("Loaded SMTP credentials from file.");
            }
        } else
        {
            Console.WriteLine("No SMTP credentials found.");
            //Ask user for SMTP Credentials
            Console.Write("Enter SMTP username: ");
            smtpUser = Console.ReadLine();
            Console.Write("Enter SMTP password: ");
            smtpPassword = Console.ReadLine();
        
            if (string.IsNullOrWhiteSpace(smtpUser) || string.IsNullOrWhiteSpace(smtpPassword))
            {
                Console.WriteLine("SMTP credentials are required.");
                Environment.Exit(1);
            }
            //Encrypt the credentials for future use.
            string encryptedSmtpUser = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(smtpUser));
            string encryptedSmtpPassword = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(smtpPassword));
            smtpUser = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(encryptedSmtpUser));
            smtpPassword = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(encryptedSmtpPassword));

            //Ask user if they want to save the credentials for future use.
            Console.Write("Do you want to save the SMTP credentials for future use? (y/n): ");
            string? saveCredentials = Console.ReadLine();
            if (saveCredentials?.ToLower() == "y")
            {
                //Here you would implement saving the encrypted credentials to a secure location.
                //For demonstration purposes, we will save the encrypted credentials to a file in the user's Documents folder.
                credentialsPath = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Documents\Automations\Orders\smtp_credentials.txt");
                File.WriteAllText(credentialsPath, $"User:{encryptedSmtpUser}\nPassword:{encryptedSmtpPassword}");
                Console.WriteLine($"SMTP credentials saved to: {credentialsPath}");
            }
        }

        //Load footer info from footer.txt if it exists, else set to empty string
        string footerPath = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Documents\Automations\Orders\footer.txt");
        string footerInfo = "<br><br>Thank you!";

        if (File.Exists(footerPath))
        {
            footerInfo = File.ReadAllText(footerPath);
            Console.WriteLine($"Loaded footer info from: {footerPath}");
        }
        else
        {
            Console.WriteLine("No footer info found. Footer file created.");
            File.WriteAllText(footerPath, footerInfo);
        }

        //New thread to check if Enter is pressed to exit the application
        new Thread(() =>
        {
            Console.WriteLine("Press [enter] to exit.");
            Console.ReadLine();
            Environment.Exit(0);
        }).Start();

        //For each PDF file already in the directory, process it, then move it to the output directory if it is not being used. Else, wait.
        foreach (var filePath in Directory.GetFiles(directoryPath, "*.pdf"))
        {
            if (string.IsNullOrWhiteSpace(smtpUser) || string.IsNullOrWhiteSpace(smtpPassword))
            {
                Console.WriteLine("SMTP credentials are required.");
                Environment.Exit(1);
            }
            ProcessFiles(filePath, smtpUser, smtpPassword);
        }


        //Every 10 seconds, check the directory for new PDF files and process them. Exit when the user presses [enter].
        while (true)
        {
            foreach (var filePath in Directory.GetFiles(directoryPath, "*.pdf"))
            {
                //Process the file (similar to the code above)
                ProcessFiles(filePath, smtpUser, smtpPassword);
            }
            Thread.Sleep(10000);
        }

        //send email with smtp with file attachment
        void send_email(string recipientEmail, string subject, string body, string[] attachmentPaths, string smtpUser, string smtpPassword)
        {
            
        
            Console.WriteLine($"Sending email to: {recipientEmail}");

            //implement the actual SMTP email sending logic using the smtpServer, smtpPort, smtpUser, and smtpPassword credentials.
            //Catch Exceptions 
            try
            {
                using (var client = new System.Net.Mail.SmtpClient(smtpServer, smtpPort))
                {
                    client.Credentials = new System.Net.NetworkCredential(smtpUser, smtpPassword);
                    client.EnableSsl = true;
                    //authenticate with the SMTP server using the provided credentials
                    client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;

                    var mailMessage = new System.Net.Mail.MailMessage();
                    
                    mailMessage.From = new System.Net.Mail.MailAddress(smtpUser);
                    mailMessage.To.Add(recipientEmail);
                    mailMessage.Subject = subject;
                    mailMessage.Body = body + footerInfo;
                    //mailMessage.IsBodyHtml = true;

                    foreach (var attachmentPath in attachmentPaths)
                    {
                        mailMessage.Attachments.Add(new System.Net.Mail.Attachment(attachmentPath));
                    }
                    client.Send(mailMessage);
                    mailMessage.Dispose();
                    Console.WriteLine($"{subject} - Email sent successfully to {recipientEmail}");
                }
                
                //Move the file to the output directory after sending the email
                string destinationPath = Path.Combine(outputDirectoryPath, Path.GetFileName(attachmentPaths[0]));
                bool fileMoved = false;
                while (!fileMoved)
                {
                    try
                    {
                        //If the file already exists in the output directory, rename the file with a number suffix and try again
                        if (File.Exists(destinationPath))
                        {
                            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(attachmentPaths[0]);
                            string extension = Path.GetExtension(attachmentPaths[0]);
                            int counter = 1;
                            do
                            {
                                destinationPath = Path.Combine(outputDirectoryPath, $"{fileNameWithoutExtension}_{counter}{extension}");
                                counter++;
                            } while (File.Exists(destinationPath));
                        }
                        
                        //Move the file to the output directory
                        File.Move(attachmentPaths[0], destinationPath);
                        Console.WriteLine($"File moved to: {destinationPath}");
                        fileMoved = true;
                        
                    }
                    catch (IOException e)
                    {
                        
                        Console.WriteLine($"Unable to move file. File {Path.GetFileName(attachmentPaths[0])} is in use. Retrying in 1 second...");
                        Console.WriteLine(e.Message);
                        Thread.Sleep(1000);
                    }
                }
            }
            //If credentials are invalid, stop the application to prevent further attempts to send emails with invalid credentials.
            //For other reasons, keep the application running to allow for retries or further processing.
            catch (Exception ex)
            {
                if (ex is System.Net.Mail.SmtpException smtpEx && smtpEx.StatusCode == System.Net.Mail.SmtpStatusCode.GeneralFailure)
                {
                    Console.WriteLine("Invalid SMTP credentials detected.");
                    Console.WriteLine("Stopping application due to email sending failure.");
                    Environment.Exit(1);
                }
                else
                {
                    Console.WriteLine($"Failed to send email to {recipientEmail}: {ex.Message}");
                    Console.WriteLine("Stopping application due to email sending failure.");
                    Environment.Exit(1);
                }
                                
            }

        }

        void ProcessFiles(string filePath, string? smtpUser, string? smtpPassword)
        {
            //Ensure there is a non-null smtpUser and smtpPassword before proceeding
            if (string.IsNullOrWhiteSpace(smtpUser) || string.IsNullOrWhiteSpace(smtpPassword))
            {
                Console.WriteLine("SMTP credentials are required.");
                Environment.Exit(1);
            }

            //Extract the patient's name from the filename, regex "^(.*?)(\d{8})", if unable, name is Unknown
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            var match = System.Text.RegularExpressions.Regex.Match(fileName, @"^(.*?)(\d{8})");
            string patientName = match.Success ? match.Groups[1].Value.Trim() : "Unknown";
            Console.WriteLine($"Processing file: {fileName}, Patient Name: {patientName}");

            //Determine the recipient email based on the file name pattern
            string recipientEmail = emailMappings.FirstOrDefault(mapping => fileName.Contains(mapping.Key)).Value;

            //If "- Test" is found in the file name, set the recipient email to a specific address
            if (fileName.Contains("- Test"))
            {
                //Do Nothing, as the recipient email is already set to the test email address in the emailMappings dictionary
            } 
            //If (ECAH) is found in the file name, set the recipient email to a specific address
            else if (fileName.Contains("ECAH"))
            {
                recipientEmail = "ecah@allcareprovider.com";
            }
            else
            {
                if (string.IsNullOrEmpty(recipientEmail))
                {
                    Console.WriteLine($"No email mapping found for file: {fileName}");
                    return;
                }
                else
                {
                    Console.WriteLine($"Recipient email determined: {recipientEmail}");
                }
            }
            //Set the subject and body for the email and send the email with the determined subject and body
            //Append "Thank you!" to the body of the email with bold formatting.
            if(fileName.Contains("ROC Order"))
            {
                string subject = $"ROC Order for {patientName}";
                string body = $"The ROC order for {patientName} is attached.";
                send_email(recipientEmail, subject, body, new string[] { filePath }, smtpUser, smtpPassword);
            }
            else if(fileName.Contains("Order -"))
            {
                string subject = $"Order for {patientName}";
                string body = $"The order for {patientName} is attached.";
                send_email(recipientEmail, subject, body, new string[] { filePath }, smtpUser, smtpPassword);
            }
            else
            {
                string subject = $"File: {patientName}";
                string body = $"The file for {patientName} is attached.";
                send_email(recipientEmail, subject, body, new string[] { filePath }, smtpUser, smtpPassword);
            }
        }
    }
}
