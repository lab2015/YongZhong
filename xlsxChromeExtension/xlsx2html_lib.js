
/*Author: Lv Shuqi  Day: 2015.09.10*/

//解析
string2xml = function(dataStr)
{
    if (!window.DOMParser && window.ActiveXObject) {
      //for IE
      xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
      xmlDoc.async="false";
      xmlDoc.loadXML(dataStr);
      return xmlDoc;
   }
   else if (document.implementation && document.implementation.createDocument) {
      //for Mozila
      parser=new DOMParser();
      xmlDoc=parser.parseFromString(dataStr,"text/xml");
      return xmlDoc;
   }
   else{
        return null;
    }
}

ch2int = function(ch)
{
    if(ch>='A' && ch<='Z')
    {
        return ch.charCodeAt()-64;
    }
    return 0;
}

int2ch = function(num)
{
    return String.fromCharCode(num+64);
}

int2str = function(num)
{
    var str = "";
    do
    {
        var md = num % 26;
        if(md == 0)
            md = 26;
        str = int2ch(md) + str;
        num = Math.floor(num/27);
    }while(num != 0)
    return str;
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
	
    var stringList = [];
    var text_content=[]; //[0][i]和[i][0]不存数据	
	(function () 
    {
        var $text_content = $("#text_content");
        // remove content
        $text_content.html("<p>");

        var f = files[0];
        var reader = new FileReader();

        reader.onload = (function(theFile) 
        {
            return function(e) 
            {
                var $htmlContent = $("<p>");
           
                try 
                {
                    // read the content of the file with JSZip
                    var zip = new JSZip(e.target.result);

                    $.each(zip.files, function (index, zipEntry) 
                    {                
                        if(zipEntry.name === "xl/sharedStrings.xml")
                        {
                            var sharedStringsDoc = string2xml(zipEntry.asText());
                            if(sharedStringsDoc == null)
                            {
                                alert("Unsupported xml doc!");
                            }
                            var nodelist = sharedStringsDoc.documentElement.childNodes;
                            var length = sharedStringsDoc.documentElement.childNodes.length;
                            stringList = [];
                            
                            for(var i = 0; i < length; i++)
                            {
                                stringList[i] = nodelist[i].childNodes[0].childNodes[0].nodeValue;
                            }
                        }
                    });
                    $.each(zip.files, function (index, zipEntry) 
                    { 
                        if(zipEntry.name == "xl/worksheets/sheet1.xml")
                        {
                            var sheetDoc = string2xml(zipEntry.asText());
                            if(sheetDoc == null)
                            {
                                alert("not support xml doc!");
                            }
                            nodelist = sheetDoc.getElementsByTagName("sheetData")[0].childNodes;
                            length = nodelist.length;
                            
                            var maxcols = 0;
                            var maxrow = 0;
                            for(var i = 0; i < length; i++)
                            {
                                var colslist = nodelist[i].childNodes;
                                colsnum = colslist.length;
                                var row = parseInt(nodelist[i].getAttribute("r"));
                                if(row > maxrow)
                                {
                                    maxrow = row;
                                }
                                for(var j = 0; j < colsnum; j++)
                                {
                                    var rowCols = colslist[j].getAttribute("r");
                                    var len = rowCols.length;
                                    var cols = 0;
                                    for(var k = len-1, l = 0; k >= 0; k--)
                                    {
                                        if(ch2int(rowCols.substr(k,1)) == 0)
                                            continue;
                                        cols += ch2int(rowCols.substr(k,1))*Math.pow(26,l);
                                        l++;
                                    }
                                    if(cols > maxcols)
                                    {
                                        maxcols = cols;
                                    }
                                }
                            }

                            text_content=[]; //[0][i]和[i][0]不存数据
                            for(var i = 0; i <= maxrow; i++)
                            {
                                text_content[i] = [];
                                for(var j = 0; j <= maxcols; j++)
                                {
                                    text_content[i][j] = null;
                                }
                            }                            
                    
                            for(var i = 0; i < length; i++)
                            {
                                var colslist = nodelist[i].childNodes;
                                colsnum = colslist.length;

                                var row = parseInt(nodelist[i].getAttribute("r"));

                                for(var j = 0; j < colsnum; j++)
                                {
                                    var rowCols = colslist[j].getAttribute("r");
                                    var len = rowCols.length;
                                    var cols = 0;

                                    for(var k = len-1, l = 0; k >= 0; k--)
                                    {
                                        if(ch2int(rowCols.substr(k,1)) == 0)
                                            continue;
                                        cols += ch2int(rowCols.substr(k,1))*Math.pow(26,l);
                                        l++;    
                                    }

                                    var value = colslist[j].childNodes[0].childNodes[0].nodeValue;

                                    if(colslist[j].hasAttribute("t"))
                                    {
                                        if(colslist[j].getAttribute("t") == "s")
                                        {
                                            text_content[row][cols] = stringList[parseInt(value)];
                                        }
                                    }
                                    else
                                    {
                                        text_content[row][cols] = value;           
                                    }                                    
                                }
                            }

                            var excelContent = "";
                            excelContent += "<table border='1px'>";
                            excelContent += "<th>  </th>";
                            for(var i = 1; i <= maxcols; i++)
                            {
                                excelContent += "<th>";
                                excelContent += int2str(i);
                                excelContent += "</th>";
                            }
                            for(var i = 1; i <= maxrow; i++)
                            {
                                excelContent += "<tr>";
                                excelContent += "<th>";
                                excelContent += i;
                                excelContent += "</th>";
                                for(var j = 1; j <= maxcols; j++)
                                {
                                    excelContent += "<td>";
                                    if(text_content[i][j]!=null)
                                    {
                                        excelContent += text_content[i][j];
                                    }
                                    excelContent += "</td>";
                                }
                                excelContent += "</tr>";
                            }
                            excelContent += "</table>";
                            $text_content.append(excelContent);
                        }
                    });
                } 
                catch(e) {}
                $text_content.append($htmlContent);
            }
        })(f);
    reader.readAsBinaryString(f);
    })();
}
window.addEventListener("load", addDNDListeners, false);