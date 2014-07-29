using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace MyWindowsService
{
  // This service will monitor "C:\FolderToMonitor" and if any ipt file
  // is added, then it will update "C:\FolderToMonitor\log.txt" with 
  // information about its iProperties
  // http://msdn.microsoft.com/en-us/library/zt39148a(v=vs.110).aspx
  public partial class MyService : ServiceBase
  {
    string folderToMonitor = @"\FolderToMonitor";
    string logFile = @"\FolderToMonitor\log.txt";
    FileSystemWatcher watcher;

    public MyService()
    {
      InitializeComponent();
    }

    protected override void OnStart(string[] args)
    {
      // Set file paths
      string documentsPath = 
        System.Environment.GetFolderPath(
          System.Environment.SpecialFolder.CommonDocuments);
      folderToMonitor = documentsPath + folderToMonitor;
      logFile = documentsPath + logFile;

      // http://msdn.microsoft.com/en-us/library/system.io.filesystemwatcher(v=vs.110).aspx
      watcher = new FileSystemWatcher();
      watcher.Path = folderToMonitor;
      watcher.Filter = "*.ipt";
      watcher.Created += watcher_Created;
      watcher.EnableRaisingEvents = true;
    }

    void watcher_Created(object sender, FileSystemEventArgs e)
    {
      // A file got created. Let's check its iProperties
      string iProperties = "File path: " + e.FullPath + "\r\n";
       
      // We might not always be able to open a document, e.g.
      // maybe if the writer locked it, so let's catch any errors
      try
      {
        ApprenticeServerComponent app = new ApprenticeServerComponent();
        ApprenticeServerDocument doc = app.Open(e.FullPath);

        // Gather "Summary Information" properties
        PropertySet ps = doc.PropertySets["{F29F85E0-4FF9-1068-AB91-08002B27B3D9}"];
        foreach (Property p in ps)
        {
          // e.g. the Thumbnail property cannot be converted to string
          // so that would throw an error we need to catch
          try
          {
            iProperties += p.DisplayName + ": " + p.Value.ToString() + "\r\n";
          }
          catch { }
        }

        app.Close();
        app = null;
      }
      catch (Exception ex)
      {
        iProperties += "Exception occurred: " + ex.Message + "\r\n";
      }

      // Write it to the file
      using (StreamWriter sw = System.IO.File.AppendText(logFile))
      {
        sw.Write(iProperties);
      }
    }

    protected override void OnStop()
    {
      watcher.EnableRaisingEvents = false;
      watcher = null;
    }
  }
}
