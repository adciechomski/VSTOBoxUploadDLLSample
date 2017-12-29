using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Runtime.InteropServices;
using RestSharp;

namespace BoxUpload
{
    [ComVisible(true)]
    public interface IBoxCalls
    {
        void Upload_Doc(string folderId, string accessToken, string filePath, string fileName);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class BoxCalls : IBoxCalls
    {
        public void Upload_Doc(string folderId, string accessToken, string filePath, string fileName)
        {
            var client = new RestClient("https://upload.box.com/api/2.0");
            var request = new RestRequest("files/content", Method.POST);
            request.AddParameter("parent_id", folderId);
            request.AddHeader("Authorization", "Bearer " + accessToken);
            string path = filePath + fileName;
            byte[] byteArray = System.IO.File.ReadAllBytes(path);
            request.AddFile("filename", byteArray, fileName);
            var responses = client.Execute(request);
            var content = responses.Content;
        }
    }
}
