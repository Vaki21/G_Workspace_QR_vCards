using Google.Apis.Auth.OAuth2;
using Google.Apis.Admin.Directory.directory_v1;
using Google.Apis.Admin.Directory.directory_v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Thought.vCards;
using QRCoder;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Net;
using Renci.SshNet;
using Renci.SshNet.Sftp;
using System.Security.Cryptography;

namespace AdminSDKDirectoryQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/admin-directory_v1-dotnet-quickstart.json
        static string[] Scopes = { DirectoryService.Scope.AdminDirectoryUserReadonly };
        static string ApplicationName = "Directory API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Directory API service.
            var service = new DirectoryService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.
            UsersResource.ListRequest request = service.Users.List();
            request.Customer = "my_customer";
            request.MaxResults = 400;
            request.OrderBy = UsersResource.ListRequest.OrderByEnum.Email;

            // For each user generates a vCard into a QR Code
            IList<User> users = request.Execute().UsersValue;
            Console.WriteLine("Users:");
            if (users != null && users.Count > 0)
            {
                //Gets this month and year
                string thisMonth = DateTime.Now.ToString("MM");
                string thisYear = DateTime.Now.ToString("yyyy");
                
                foreach (var userItem in users)
                {
                    Console.WriteLine("{0} ({1}) [{2}]", userItem.PrimaryEmail,
                        userItem.Name.GivenName, userItem.RecoveryPhone);

                    // Generates vCard
                    vCard vCard = new vCard();
                    //URL is private so couldn't make it work
                    //Console.WriteLine(userItem.ThumbnailPhotoUrl);
                    //if (userItem.ThumbnailPhotoUrl != null) { vCard.Photos.Add( new vCardPhoto(userItem.ThumbnailPhotoUrl) ); }
                    vCard.GivenName = userItem.Name.GivenName;
                    vCard.FamilyName = userItem.Name.FamilyName;

                    /*
                    //Assigns picture from Google DB, links by google are private so I don't how to make this work
                    HttpWebRequest myHttpWebRequest = (HttpWebRequest)WebRequest.Create(userItem.ThumbnailPhotoUrl);
                    myHttpWebRequest.AllowAutoRedirect = true;
                    HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
                    string redirUrl = myHttpWebResponse.ResponseUri.ToString();
                    myHttpWebResponse.Close();
                    Console.WriteLine(redirUrl);

                    using (WebClient webClient = new WebClient())
                    {
                        webClient.DownloadFile(redirUrl, userItem.PrimaryEmail + "_photo.jpeg");
                    }

                    string photo = userItem.PrimaryEmail + "_photo.jpeg";
                    vCard.Photos.Add(new vCardPhoto(photo));
                    */

                    //Assigns Job Title if user has one
                    if (userItem.Organizations != null) { vCard.Title = userItem.Organizations[0].Title; }

                    vCard.EmailAddresses.Add(new vCardEmailAddress(userItem.PrimaryEmail, vCardEmailAddressType.Internet));
                    vCard.Organization = "";
                    vCard.Websites.Add(new vCardWebsite(""));

                    //Assign a number if user has one
                    if (userItem.Phones != null) { vCard.Phones.Add(new vCardPhone(userItem.Phones[0].Value)); }

                    vCardDeliveryAddress address = new vCardDeliveryAddress();
                    address.Street = "";
                    address.City = "Praha";
                    //address.PostalCode = "";
                    //address.Country = "Slovakia";

                    vCardDeliveryAddress address1 = new vCardDeliveryAddress();
                    address1.Street = "";
                    address1.City = "Praha";
                    //address1.PostalCode = "";
                    //address1.Country = "Czechia";

                    //Assigns Address based on Phone number Country code
                    if (userItem.Phones != null)
                    {

                        //Slovakia Address
                        if (userItem.Phones[0].Value.Contains("+421"))
                        {
                            vCard.DeliveryAddresses.Add(address);
                        }
                        //Czech Address
                        else if (userItem.Phones[0].Value.Contains("+420"))
                        {
                            vCard.DeliveryAddresses.Add(address1);
                        }
                        //If User doesnt have Czech/Slovak country code, assign both
                        else
                        {
                            vCard.DeliveryAddresses.Add(address);
                            vCard.DeliveryAddresses.Add(address1);
                        }
                    }
                    //If user doesnt have a phone number, assign both addresses
                    else
                    {
                        vCard.DeliveryAddresses.Add(address);
                        vCard.DeliveryAddresses.Add(address1);
                    }
                    
                    /*
                    // Takes address from Google Admin, wasn't worth it in my case
                    if (userItem.Addresses != null)
                    {
                        vCardDeliveryAddress address = new vCardDeliveryAddress();
                        address.Street = userItem.Addresses[0].Formatted;
                        vCard.DeliveryAddresses.Add(address);
                    }
                    */

                    
                    // Save vCard data to string
                    vCardStandardWriter writer = new vCardStandardWriter();
                    StringWriter stringWriter = new StringWriter();
                    writer.Write(vCard, stringWriter);

                    //Console.WriteLine(stringWriter.ToString());


                    // Creates QR Codes and saves them into BMP
                    QRCodeGenerator qrGenerator = new QRCodeGenerator();
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode(stringWriter.ToString(), QRCodeGenerator.ECCLevel.L);
                    QRCode qrCode = new QRCode(qrCodeData);
                    Bitmap qrCodeImage = qrCode.GetGraphic(2, Color.Black, Color.White, true);
                    //Adds logo in the QR Code
                    //Bitmap qrCodeImage = qrCode.GetGraphic(2, Color.Black, Color.White, (Bitmap)Bitmap.FromFile("logo.png"), 20,0,true);
                    if (Directory.Exists("QRImages"))
                    {
                        qrCodeImage.Save("QRImages\\" + userItem.PrimaryEmail + "_border.bmp");
                    }
                    else
                    {
                        Directory.CreateDirectory("QRImages");
                        qrCodeImage.Save("QRImages\\" + userItem.PrimaryEmail + "_border.bmp");
                    }

                    //Location of generated QR Code
                    string oFileName = "QRImages\\" + userItem.PrimaryEmail + "_border.bmp";
                    //Load image from file and crop border it into original position
                    using (Image codeQR = Image.FromFile(oFileName))
                    {
                        var croppedWidth = codeQR.Width - ((codeQR.Width / 100) * 14);
                        var croppedHeight = codeQR.Height - ((codeQR.Height / 100) * 14);
                        var croppedPos = ((codeQR.Width / 100) * 14) / 2;

                        // Create a new image at the cropped size
                        Bitmap cropped = new Bitmap(croppedWidth, croppedHeight);

                        // Create a Graphics object to do the drawing, with the new bitmap as the target
                        using (Graphics g = Graphics.FromImage(cropped))
                        {
                            // Draw the desired area of the original into the graphics object
                            g.DrawImage(codeQR, new Rectangle(0, 0, croppedWidth, croppedHeight), new Rectangle(croppedPos, croppedPos, croppedWidth, croppedHeight), GraphicsUnit.Pixel);
                            string fileName = "QRImages\\" + userItem.PrimaryEmail + "_cropped_" + thisMonth + ".bmp";
                            // Save the result
                            cropped.Save(fileName);
                            cropped.Dispose();
                        }
                        codeQR.Dispose();
                    }
                    //Delete Original QR Code
                    File.Delete(oFileName);
                }
                Console.WriteLine("QR Codes Generated...");

                //--------------------Establishes SSH connection and uploads QR files to the server---------------------------------
                const string host = "";
                const string username = "";
                //var privateKey = new PrivateKeyFile(@"C:\\Users\\XXXX\\.ssh\\id_rsa");
                var privateKey = new PrivateKeyFile(@"id_rsa");

                const string workingDirectory = "";
                string newDirectory = workingDirectory + "/" + thisYear + "/" + thisMonth;
                //var localDirectory = Directory.GetFiles("C:\\Users\\XXXX\\source\\repos\\QRUsers\\QRUsers\\bin\\Debug\\net6.0\\QRImages");
                var localDirectory = Directory.GetFiles("QRImages");
                
                using (var client = new SftpClient(host, username, new[] { privateKey }))
                {
                    client.Connect();
                    Console.WriteLine("Connected to {0}", host);
                    
                    //Lists working directory
                    client.ChangeDirectory(workingDirectory);
                    Console.WriteLine("Changed directory to {0}", workingDirectory);
                    var listDirectory = client.ListDirectory(workingDirectory);
                    Console.WriteLine("Listing directory:");
                    foreach (var files in listDirectory)
                    {
                        Console.WriteLine(" - " + files);
                        //var kok = System.IO.Path.GetFileName(files);
                        //client.DeleteFile("/wp-content/uploads/test/" + kok);
                    }
                    //Checks if new directory exists, if not it creates it
                    if (!(client.Exists(workingDirectory + "/" + thisYear)))
                    {
                        client.CreateDirectory(workingDirectory + "/" + thisYear);
                        client.ChangeDirectory(workingDirectory + "/" + thisYear);
                        if (!(client.Exists(newDirectory)))
                        {
                            client.CreateDirectory(newDirectory);
                        }
                    }
                    else
                    {
                        if (!(client.Exists(newDirectory)))
                        {
                            client.CreateDirectory(newDirectory);
                        }
                    }
                    //Changes directory to the new one
                    client.ChangeDirectory(newDirectory);

                    //Checks if the file already exists or not and starts to upload it
                    foreach (var file in localDirectory)
                    {
                        using (var fileStream = new FileStream(file, FileMode.Open))
                        {
                            if (!(client.Exists(newDirectory + "/" + Path.GetFileName(file))))
                            {
                                Console.WriteLine("Uploading {0} ({1:N0} bytes)", file, fileStream.Length);
                                client.BufferSize = 4 * 1024; // bypass Payload error large files
                                client.UploadFile(fileStream, Path.GetFileName(file));
                                fileStream.Dispose();
                            }
                            else
                            {
                                Console.WriteLine("File already exists");
                                fileStream.Dispose();
                            }
                        }
                    }
                    //Shows the uploaded result and disconnects
                    Console.WriteLine("Changed directory to {0}", newDirectory);
                    var listDirectory2 = client.ListDirectory(newDirectory);
                    Console.WriteLine("Listing directory:");
                    foreach (var files in listDirectory2)
                    {
                        Console.WriteLine(" - " + files.Name);
                    }
                    Console.WriteLine("Files Uploaded and client disconnected...");
                }
            }
            else
            {
                Console.WriteLine("No users found...");
            }
            Console.Read();
        }
    }
}