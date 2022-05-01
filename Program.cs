using System.Collections.Generic;

namespace Json
{
    using System;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Runtime.Serialization.Json;
    using System.Text.RegularExpressions;
    using System.Xml;
    using System.Xml.Linq;
    using System.Xml.Serialization;
    using Newtonsoft.Json.Linq;
    using Newtonsoft.Json;
	using ClosedXML.Excel;
	
    class Program
    {
        static void Main(string[] args)
        {
            var converter = new Converter();
            converter.Convert();
			var loadXml = new LoadXml();
			loadXml.load();

        }
    }
  
  class LoadXml
  {
	  
    public void load()
    {
    var xmlData = new List<DependenciesXml>();		
  foreach (var file in 
   Directory.EnumerateFiles(Environment.ExpandEnvironmentVariables("%userprofile%\\downloads"), "*.xml"))
  { 

 //Console.WriteLine(file);
 try
	{
		var xml = XElement.Load(file);

		var listToIgnore = new[] { "np*.", "npsharp*.", "system*.", "microsoft*." };

		var versionList = xml.Descendants("file")
							.Where(_ => !Regex.IsMatch(_.Attribute("name").Value, $"^({ string.Join("|", listToIgnore) })", RegexOptions.IgnoreCase))
							.Select(_ => new XmlFileVersion { Name = _.Attribute("name")?.Value, Version = _.Attribute("version")?.Value })
							.Distinct(new XmlFileVersionComparer())
							.OrderBy(_ => _.Name);
      
	  //Create table of xlsx file
	 DataTable Table = new DataTable();
	 Table.Clear(); 
	 Table.Columns.Add("Name");
     Table.Columns.Add("Version");
	 
	 //Read data of xml file
	 foreach (var entry in versionList) {
	  DataRow dr = Table.NewRow();
	  var getDataXml = new DependenciesXml() {Name = entry.Name,Version = entry.Version};
      xmlData.Add(getDataXml);
	 }
	 
	 //Add data in the table
	 foreach (var outputData in xmlData) {
	  DataRow dr = Table.NewRow();
	  dr["Name"] = outputData.Name;
      dr["Version"] = outputData.Version;
      Table.Rows.Add(dr);
	 }
	 
	 //Build xlsx file 
	 XLWorkbook wb = new XLWorkbook();
     wb.Worksheets.Add(Table, "Subnet Data");
     wb.SaveAs(Environment.ExpandEnvironmentVariables("%userprofile%/downloads/HTMLDependencies.xlsx"));
	 
	}
	catch (FileNotFoundException)
	{
		Console.WriteLine("File not found. Try downloading ExtendedBuildSpecification.xml from TeamCity to your Downloads folder and run again");
		throw;
	}
}

}

}

class XmlFileVersionComparer : IEqualityComparer<XmlFileVersion>
{
	public bool Equals(XmlFileVersion x, XmlFileVersion y)
	{
		if (Object.ReferenceEquals(x, y)) return true;
		if (Object.ReferenceEquals(x, null) || Object.ReferenceEquals(y, null))
			return false;

		return x.Name == y.Name && x.Version == y.Version;
	}

	public int GetHashCode(XmlFileVersion fileVersion)
	{
		if (Object.ReferenceEquals(fileVersion, null)) return 0;
		var hashName = fileVersion.Name == null ? 0 : fileVersion.Name.GetHashCode();
		var hashVersion = fileVersion.Version == null ? 0 : fileVersion.Version.GetHashCode();

		return hashName ^ hashVersion;
	}
}

class XmlFileVersion
{
	public string Name { get; set; }
	public string Version { get; set; }
}


    class Converter
    {
        public void Convert()
        {
            //  your JSON string goes here
            string filePath = Environment.ExpandEnvironmentVariables("%userprofile%/downloads/packages.json");
            string jsonString = File.ReadAllText(filePath);
            
            // deconstruct the JSON
            var root = new FileList { Packages = new List<devDependencies>() };
            var jObject = JObject.Parse(jsonString);

             foreach (var devDependencieMessageJsonObject in jObject["devDependencies"])
              {
               var jProperty = (JProperty)devDependencieMessageJsonObject;
               var devDependencieDescription = jProperty.Name.ToString();
               var devDependencieCode =  jProperty.Value.ToString();
               var devDependencieMessage = new devDependencies() {nameAttr = devDependencieDescription,
                                                                  name = devDependencieDescription,
                                                                  versionAttr = devDependencieCode, 
                                                                  version = devDependencieCode};
               root.Packages.Add(devDependencieMessage);
             }
             
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.IndentChars = ("\t");
            settings.OmitXmlDeclaration = true;
            var xmlSerializer = new XmlSerializer(typeof(FileList));
            string xml;
            using (XmlWriter textWriter = XmlWriter.Create(Environment.ExpandEnvironmentVariables("%userprofile%/downloads/HTMLDependencies.xml"), settings))
            {
                xmlSerializer.Serialize(textWriter, root);
                xml = textWriter.ToString();
            }
            Console.WriteLine("The file was created !");
        }
			   
    }
    
    [XmlRoot("Packages")]
    public class FileList
    {
        [XmlElement("file")]
        public List<devDependencies> Packages {get;set;}
    }

    public class devDependencies
    {  
    
     [XmlAttribute("name")]
     public String nameAttr { get; set; }
     public String name { get; set; }
     [XmlAttribute("version")]
     public String versionAttr { get; set; }
     public String version { get; set; }
         
    }
	
	 public class AllXmlList
    {
        public List<DependenciesXml> DataXml {get;set;}
    }
	
	public class DependenciesXml
    {  
  
     public string Name { get; set; }
     public string Version { get; set; }
         
    }
		       
}