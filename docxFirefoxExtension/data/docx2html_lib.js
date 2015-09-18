
/*Author: Geng Si  Day: 2015.09.10*/

function loadXMLStr(xmlString){  
    var xmlDoc=null;  
    // If IE browser
    if(!window.DOMParser && window.ActiveXObject){   // IE browser  
        var xmlDomVersions = ['MSXML.2.DOMDocument.6.0','MSXML.2.DOMDocument.3.0','Microsoft.XMLDOM'];  
        for(var i=0;i<xmlDomVersions.length;i++){  
            try{  
                xmlDoc = new ActiveXObject(xmlDomVersions[i]);  
                xmlDoc.async = false;  
                xmlDoc.loadXML(xmlString);   
                break;  
            }catch(e){  
            }  
        }  
    }  
    // If Mozilla browser  
    else if(window.DOMParser && document.implementation && document.implementation.createDocument)
    {  
        try
        {  
            domParser = new  DOMParser();  
            xmlDoc = domParser.parseFromString(xmlString, 'text/xml');  
        }catch(e){}  
    }  
    else
    {  
        return null;  
    }  
    return xmlDoc;  
}  
     
function addDNDListeners()
{
    var container = document.getElementById("container");
    container.addEventListener("dragenter", function(event)
    {
        event.stopPropagation();
        event.preventDefault();
    }, false);
    container.addEventListener("dragover", function(event)
    {
        event.stopPropagation();
        event.preventDefault();
    }, false);
	container.addEventListener("drop", handleDrop, false);
}

function handleDrop(event)
{
    var files = event.dataTransfer.files;
	event.stopPropagation();
	event.preventDefault();
		
	(function () 
    {
        var $text_content = $("#text_content");

        // remove content
        $text_content.html("");

        var f = files[0];
        var reader = new FileReader();

        reader.onload = (function(theFile) 
        {
            return function(e) 
            {
                var $textContent = $("");
                    
                try 
                {
                    // read the content of the file with JSZip
                    var zip = new JSZip(e.target.result);

                    $.each(zip.files, function (index, zipEntry) 
                    {               
                        if(zipEntry.name === "word/document.xml")
                        {
                            var xmldoc=loadXMLStr(zipEntry.asText());
                            var p_elements = xmldoc.getElementsByTagNameNS("http://purl.oclc.org/ooxml/wordprocessingml/main","p");
                            if (p_elements.length <= 0) // Old version
                            {
                                p_elements = xmldoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main","p");

                            }
                            for (var i = 0; i < p_elements.length; i++) 
                            {  
                                var t_elements = p_elements[i].getElementsByTagNameNS("http://purl.oclc.org/ooxml/wordprocessingml/main","t");
                                if (t_elements.length <= 0) // Old version
                                {
                                    t_elements = p_elements[i].getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main","t");

                                }
                                $textContent = $("<p>");
                                for (var j = 0; j < t_elements.length; j++) 
                                { 
                                    var value = t_elements[j].firstChild.nodeValue;
                                    $textContent.append($("<span>", {text : value}));
                                }
                                $text_content.append($textContent);
                            }  
                        }
                    });
                } 
                catch(e) {}               
            }
         })(f);
    reader.readAsBinaryString(f);
    })();

}
window.addEventListener("load", addDNDListeners, false);