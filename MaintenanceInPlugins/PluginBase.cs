﻿using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using System.IO;
using System.Reflection;

namespace MaintenanceIn.MaintenanceInPlugins
{
    public abstract class PluginBase : IPlugin
    {
        /// <summary>
        /// Plug-in context object. 
        /// </summary>
        protected class LocalPluginContext
        {
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "LocalPluginContext")]
            internal IServiceProvider ServiceProvider { get; private set; }

            /// <summary>
            /// The Microsoft Dynamics 365 organization service.
            /// </summary>
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "LocalPluginContext")]
            internal IOrganizationService OrganizationService { get; private set; }

            /// <summary>
            /// IPluginExecutionContext contains information that describes the run-time environment in which the plug-in executes, information related to the execution pipeline, and entity business information.
            /// </summary>
            internal IPluginExecutionContext PluginExecutionContext { get; private set; }

            /// <summary>
            /// Synchronous registered plug-ins can post the execution context to the Microsoft Azure Service Bus. <br/> 
            /// It is through this notification service that synchronous plug-ins can send brokered messages to the Microsoft Azure Service Bus.
            /// </summary>
            internal IServiceEndpointNotificationService NotificationService { get; private set; }

            /// <summary>
            /// Provides logging run-time trace information for plug-ins. 
            /// </summary>
            internal ITracingService TracingService { get; private set; }

            private LocalPluginContext() { }

            /// <summary>
            /// Helper object that stores the services available in this plug-in.
            /// </summary>
            /// <param name="serviceProvider"></param>
            internal LocalPluginContext(IServiceProvider serviceProvider)
            {
                if (serviceProvider == null)
                {
                    throw new InvalidPluginExecutionException("serviceProvider");
                }

                // Obtain the execution context service from the service provider.
                PluginExecutionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

                // Obtain the tracing service from the service provider.
                TracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

                // Get the notification service from the service provider.
                NotificationService = (IServiceEndpointNotificationService)serviceProvider.GetService(typeof(IServiceEndpointNotificationService));

                // Obtain the organization factory service from the service provider.
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

                // Use the factory to generate the organization service.
                OrganizationService = factory.CreateOrganizationService(PluginExecutionContext.UserId);
            }

            /// <summary>
            /// Writes a trace message to the CRM trace log.
            /// </summary>
            /// <param name="message">Message name to trace.</param>
            internal void Trace(string message)
            {
                if (string.IsNullOrWhiteSpace(message) || TracingService == null)
                {
                    return;
                }

                if (PluginExecutionContext == null)
                {
                    TracingService.Trace(message);
                }
                else
                {
                    TracingService.Trace(
                        "{0}, Correlation Id: {1}, Initiating User: {2}",
                        message,
                        PluginExecutionContext.CorrelationId,
                        PluginExecutionContext.InitiatingUserId);
                }
            }
        }

        /// <summary>
        /// Gets or sets the name of the child class.
        /// </summary>
        /// <value>The name of the child class.</value>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "PluginBase")]
        protected string ChildClassName { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="PluginBase"/> class.
        /// </summary>
        /// <param name="childClassName">The <see cref=" cred="Type"/> of the derived class.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "PluginBase")]
        internal PluginBase(Type childClassName)
        {
            try
            {
                AppDomain.CurrentDomain.AssemblyResolve += (sender, ars) =>
                {
                    var resourceName = "MaintenanceIn.MaintenanceInPlugins.Resources." + new AssemblyName(ars.Name).Name + ".dll";

                    using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                    {
                        Byte[] assemblyData = new Byte[stream.Length];
                        stream.Read(assemblyData, 0, assemblyData.Length);
                        return Assembly.Load(assemblyData);
                    }
                };
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(OperationStatus.Failed, "Error on Load Assembly!" + ex.Message);
            }
            ChildClassName = childClassName.ToString();
        }

        /// <summary>
        /// Main entry point for he business logic that the plug-in is to execute.
        /// </summary>
        /// <param name="serviceProvider">The service provider.</param>
        /// <remarks>
        /// For improved performance, Microsoft Dynamics 365 caches plug-in instances. 
        /// The plug-in's Execute method should be written to be stateless as the constructor 
        /// is not called for every invocation of the plug-in. Also, multiple system threads 
        /// could execute the plug-in at the same time. All per invocation state information 
        /// is stored in the context. This means that you should not use global variables in plug-ins.
        /// </remarks>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "CrmVSSolution411.NewProj.PluginBase+LocalPluginContext.Trace(System.String)", Justification = "Execute")]
        public void Execute(IServiceProvider serviceProvider)
        {
            if (serviceProvider == null)
            {
                throw new InvalidPluginExecutionException("serviceProvider");
            }

            // Construct the local plug-in context.
            LocalPluginContext localcontext = new LocalPluginContext(serviceProvider);

            localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Entered {0}.Execute()", this.ChildClassName));

            try
            {
                // Invoke the custom implementation 
                ExecuteCrmPlugin(localcontext);
                // now exit - if the derived plug-in has incorrectly registered overlapping event registrations,
                // guard against multiple executions.
                return;
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Exception: {0}", e.ToString()));

                // Handle the exception.
                throw new InvalidPluginExecutionException("OrganizationServiceFault", e);
            }
            finally
            {
                localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Exiting {0}.Execute()", this.ChildClassName));
            }
        }

        /// <summary>
        /// Placeholder for a custom plug-in implementation. 
        /// </summary>
        /// <param name="localcontext">Context for the current plug-in.</param>
        protected virtual void ExecuteCrmPlugin(LocalPluginContext localcontext)
        {
            // Do nothing. 
        }

        #region Custom MessageBox
        /// <summary>
        /// 
        /// </summary>
        /// <param name="operationstatus"></param>
        /// <param name="message">متن پیام</param>
        /// <param name="objectname">this</param>
        /// <param name="position"></param>
        /// <returns></returns>
        public static InvalidPluginExecutionException MessageBox(OperationStatus operationstatus, string message, Object objectname = null, double position = -1)
        {
            Type type = null;
            if (objectname != null)
                type = objectname.GetType();
            switch (operationstatus)
            {
                case OperationStatus.Failed:
                    var messageBoxFailed = new InvalidPluginExecutionException(operationstatus, string.Format("{2} - Position : {0} \nERROR : \t{1}\n\n\n\n\n\n", position, message, type == null ? "" : type.Name));
                    return messageBoxFailed;
                case OperationStatus.Canceled:
                    var messageBoxCanceled = new InvalidPluginExecutionException(operationstatus, string.Format("{1} {0}", message + "\n\n\n\n\n\n", type == null ? "" : type.Name + "\n"));
                    return messageBoxCanceled;
                case OperationStatus.Retry:
                    var messageBoxRetry = new InvalidPluginExecutionException(operationstatus, message + "\n\n\n\n\n\n");
                    return messageBoxRetry;
                case OperationStatus.Suspended:
                    var messageBoxSuspended = new InvalidPluginExecutionException(operationstatus, message + "\n\n\n\n\n\n");
                    return messageBoxSuspended;
                case OperationStatus.Succeeded:
                    var messageBoxSucceeded = new InvalidPluginExecutionException(operationstatus, message + "\n\n\n\n\n\n");
                    return messageBoxSucceeded;
                default:
                    var messageBox = new InvalidPluginExecutionException(operationstatus, string.Format("{2} - Position : {0} \nERROR : \t{1}\n\n\n\n\n\n", position, message, type == null ? "" : type.Name));
                    return messageBox;
            }
        }
        #endregion

        /// <summary>
        /// convert to shamsi date
        /// </summary>
        /// <param name="date">تاریخ میلادی</param>
        /// <returns>تاریخ شمسی</returns>
        public static string Convert2Persian(DateTime date, string separator)
        {
            PersianCalendar perdate = new PersianCalendar();
            return perdate.GetYear(date).ToString("00") + separator + perdate.GetMonth(date).ToString("00") + separator + perdate.GetDayOfMonth(date).ToString("00");
        }

        /// <summary>
        /// Merge two paths
        /// </summary>
        /// <param name="targetPath"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string CombineFileToDirectory(string targetPath, string fileName)
        {
            string temp = System.IO.Path.Combine(targetPath, fileName);
            return temp;
        }

        /// <summary>
        /// Check the file in the folder
        /// </summary>
        /// <param name="subPath"></param>
        /// <returns></returns>
        public static bool CheckFileExists(string subPath)
        {
            try
            {
                bool temp = System.IO.File.Exists(subPath);
                return temp;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// Create directory
        /// </summary>
        /// <param name="subPath"></param>
        /// <returns>برگشت مسیر ایجاد شده</returns>
        public static DirectoryInfo CreateDirectory(string subPath)
        {
            DirectoryInfo temp = System.IO.Directory.CreateDirectory(subPath);
            return temp;
        }

        /// <summary>
        /// Check the folder path
        /// </summary>
        /// <param name="subPath"></param>
        /// <returns></returns>
        public static bool CheckDirectoryExists(string subPath)
        {
            try
            {
                bool temp = System.IO.Directory.Exists(subPath);
                return temp;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

    }
}