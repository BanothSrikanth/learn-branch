public class DocxController {
    // DOCX file path inside the ZIP
    public string docxml = 'word/document.xml';
    public Set<String> filestoProcess=new Set<String>{'word/document.xml'};

    public void generateDocument(String docId) {
        List<ContentVersion> docxFile = [SELECT Id, PathOnClient,Title, VersionData FROM ContentVersion WHERE ContentDocumentId = :docId];
        if (docxFile.isEmpty()) {
            return;
        }
        
        // Read the DOCX (zipped file)
        Blob filedata = docxFile[0].VersionData;
        Compression.ZipReader ziputility = new Compression.ZipReader(filedata);
        
        //compression zip.writer to generate a new file
        Compression.ZipWriter zipWriter = new Compression.ZipWriter();
        
        for (Compression.ZipEntry entry : ziputility.getEntries()) {
            System.debug('ZipEntry==>'+entry.getName());
            if (filestoProcess.contains(entry.getName())) {
                String processedXML=processXML(ziputility.extract(entry).toString());
                zipwriter.addEntry(entry.getName(),Blob.valueOf(processedXML));
                
            }
            else if(entry.getName().contains('header')){
                string data=processHeader(ziputility.extract(entry).toString());
                zipwriter.addEntry(entry.getName(),Blob.valueOf(data));
            }
            else{
                zipWriter.addEntry(entry.getName(),ziputility.extract(entry));
            }
        }
        saveUpdatedDocx(docxFile[0].Title,zipwriter.getArchive());
    }
    
    public String processHeader(String docXml){
        return docXml.replace('REPLACE IN APEX','MERGED THE WATERMARK FROM APEX');
    }

    public String processXML(String docxml) {
        Set<String> placeholders = new Set<String>();
        Pattern pattern = Pattern.compile('\\{\\{([a-zA-Z0-9_\\.]+)\\}\\}');
        Matcher matcher = pattern.matcher(docxml);
        
        while (matcher.find()) {
            placeholders.add(matcher.group(1)); // Extract placeholders
        }
		System.debug(placeholders);
        // Fetch a list of Opportunities
        List<Opportunity> allopportunity = [SELECT Id, Name FROM Opportunity LIMIT 10];

     

        // Replace placeholders with random opportunity names
        for (String phval : placeholders) {
            Integer index = Math.floor(Math.random() * allopportunity.size()).intValue();
            docxml = docxml.replace('{{' + phval + '}}', allopportunity[index].Name);
        }
		
        return docxml;
        //System.debug('Processed XML: ' + docxml);
    }
    private void saveUpdatedDocx(String originalFileName, Blob newDocxFile) {
        ContentVersion newVersion = new ContentVersion();
        newVersion.Title = 'Updated_' + originalFileName+'.docx';
        newVersion.PathOnClient = originalFileName+'.docx';
        newVersion.VersionData = newDocxFile;
        newVersion.IsMajorVersion = true;
        insert newVersion;

        System.debug('Updated DOCX saved as new ContentVersion: ' + newVersion.Id);
    }
    
    
    
    
    
}
