
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.IO;
using Aspose.Words;
using System.Diagnostics;

namespace Docx2PDFWF
{
    public class Docx2PDFWF : CodeActivity
    {
        [Input("Prev Subject Suffix")]
        public InArgument<string> PrevSubjectSuffix { get; set; }

        [Input("New Subject Suffix")]
        public InArgument<string> NewSubjectSuffix { get; set; }


        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            ahLog("Docx2PDF Started.");
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                var FileBody = entity.Attributes["documentbody"];
                string MimeType = (string)entity.Attributes["mimetype"];
                string FileNmae = (string)entity.Attributes["filename"];
                string TempFileName = Guid.NewGuid().ToString();
                string TempPath = System.IO.Path.GetTempPath();
                //ahLog("Path: " + TempPath + "\r\n Name: " + TempFileName + "\r\nGUID:" + entity.Id.ToString() + "\r\nText:" + entity["notetext"] + "\r\nFile Name:" + entity["filename"]);
                //if (MimeType == @"application/msword" || MimeType == @"application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                if (FileNmae.ToLower().EndsWith(".docx") || FileNmae.ToLower().EndsWith(".doc"))
                {
                    try
                    {
                        using (FileStream fileStream = new FileStream(TempPath + @"\" + TempFileName, FileMode.OpenOrCreate))
                        {
                            byte[] fileContent = Convert.FromBase64String(entity["documentbody"].ToString());
                            fileStream.Write(fileContent, 0, fileContent.Length);
                        }

                        LicenseHelper.ModifyInMemory.ActivateMemoryPatching();

                        Document doc = new Document(TempPath + @"\" + TempFileName);
                        doc.Save(TempPath + @"\" + TempFileName + ".pdf", SaveFormat.Pdf);

                        Entity NewNote = new Entity("annotation");
                        NewNote["filename"] = entity["filename"] + ".pdf";
                        NewNote["documentbody"] = Convert.ToBase64String(File.ReadAllBytes(TempPath + @"\" + TempFileName + ".pdf"));
                        if (entity.Attributes.Keys.Contains("notetext"))
                            if (entity["notetext"]!=null)
                                    NewNote["notetext"] = entity["notetext"];
                        string subject = "";
                        if (entity.Attributes.Keys.Contains("subject"))
                            if (entity["subject"] != null)
                                subject = entity["subject"].ToString();
                        NewNote["subject"] = subject + "-" + NewSubjectSuffix.Get(executionContext);
                        var er = entity["objectid"] as EntityReference; ;
                        NewNote["objectid"] = new EntityReference(er.LogicalName, er.Id);

                        service.Create(NewNote);

                        ahLog("New note created successfully.");

                        if (PrevSubjectSuffix!=null)
                            if (PrevSubjectSuffix.ToString().Trim()!="")
                            {
                                subject = "";
                                if (entity.Attributes.Keys.Contains("subject"))
                                    if (entity["subject"] != null)
                                        subject = entity["subject"].ToString();
                                entity["subject"] = subject + "-" + PrevSubjectSuffix.Get(executionContext);
                                service.Update(entity);
                            }
                    }
                    catch (Exception ex)
                    {
                        ahLog(ex.ToString() + "\r\nSource:" + ex.Source.ToString());
                    }
                }
            }




        }


        private void ahLog(string Text)
        {
            EventLog.WriteEntry("Application", Text, EventLogEntryType.Error);
            /*
             * using (EventLog eventLog = new EventLog("Application"))
            {
                eventLog.Source = "Application";
                eventLog.WriteEntry(Text, EventLogEntryType.Error);
            }
            */
        }
    }
}
