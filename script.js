document.getElementById('buttonA').addEventListener('click', function () { processDocument('A'); });
document.getElementById('buttonB').addEventListener('click', function () { processDocument('B'); });
document.getElementById('buttonC').addEventListener('click', function () { processDocument('C'); });


var reportNode = "";//a variable that holds a template 'htm' report.  When the user exports the contents of the specification this is used to help generate a full Arup Report - currently there is an interim step to follow as described in the aA value report - hopefully a timesaver nonetheless

// when page loads, call this function to load the specification
window.onload = function () {
    // loadSpecification();
    fetchStub("reportStub.txt", function (fileStub) {
        reportNode = fileStub;
    })
}

//used to fetch a .htm template into which the body of the assembled specification will be copied and then through the saveHTM function served to the users downloads area
function fetchStub(url, callback) {
    var xhr = new XMLHttpRequest();
    xhr.open('GET', url, true);
    xhr.onreadystatechange = function (evt) {
        //Do not explicitly handle errors, those should be
        //visible via console output in the browser.
        if (xhr.readyState === 4) {
            callback(xhr.responseText);
        }
    };
    xhr.send(null);
}

// function processDocument(group) {
//     console.log("Processing document... " + group)
//     fetch('automateFile.docx')
//         .then(response => response.blob())
//         .then(blob => {
//             mammoth.convertToHtml({ arrayBuffer: blob.arrayBuffer() })
//                 .then(output => {
//                     let html = output.value;
//                     let modifiedHtml = removeGroup(html, group);
//                     saveDocument(modifiedHtml);
//                     saveHTM(modifiedHtml)
//                 });
//         });
// }


function processDocument(group) {
    console.log("Processing document for group: " + group);
    // fetch('automateFile.docx')
    fetch('Arup Standard Fire Engineering Scope_Property.docx')
        .then(response => response.blob())
        .then(blob => {
            // First modify the .docx file directly to remove comments
            return modifyDocxCommentsAndText(blob, group);  // This returns a promise with the modified .docx blob
        })
        .then(modifiedDocxBlob => {
            // Define style mappings to preserve specific styles
            const options = {
                styleMap: [
                    "p[style-name='Heading 1'] => h1:fresh",
                    "p[style-name='Heading 2'] => h2:fresh",
                    "p[style-name='Bullet List'] => ul > li:fresh",
                    "p[style-name='Numbered List'] => ol > li:fresh",
                    // Add mappings for alphabet-based lists, colors, and more
                    "p[style-name='Heading 3'] => ol[type='a'] > li:fresh",
                    // Use "span[style-name='ColorName']" to map specific colors
                    "r[style-name='Red Text'] => span.color-red",
                    "r[style-name='Italic Text'] => em"

                ],
                convertImage: mammoth.images.inline, // Preserve images inline
                ignoreEmptyParagraphs: true, // Ignore empty paragraphs
            };
            // Convert the modified .docx file to HTML with the style map
            return mammoth.convertToHtml({ arrayBuffer: modifiedDocxBlob.arrayBuffer() }, options);
        })
        .then(output => {
            // Process the HTML to remove additional content as needed
            let html = output.value;
            html = html.replace(/â€¯â€‹Â|â€¯â€¯/g, ''); // Remove encoding artifacts

            console.log("HTML: " + html)
            let modifiedHtml = removeGroup(html, group);
            // saveDocument(modifiedHtml);  // Assuming this saves the HTML for download
            saveHTM(modifiedHtml);  // Assuming this saves the .htm file for download
        })
        .catch(error => {
            console.error("Error during document processing:", error);
        });
}

function modifyDocxCommentsAndText(blob, group) {
    return JSZip.loadAsync(blob)
        .then(zip => {
            // Handle both comments.xml and document.xml simultaneously
            const commentsPromise = zip.file("word/comments.xml").async("string");
            const documentPromise = zip.file("word/document.xml").async("string");

            return Promise.all([commentsPromise, documentPromise]).then(values => {
                const [commentsXmlStr, documentXmlStr] = values;
                // console.log("Comments: " + commentsXmlStr)
                // console.log("Document: " + documentXmlStr)
                let parser = new DOMParser();
                let commentsDoc = parser.parseFromString(commentsXmlStr, "application/xml");
                let documentDoc = parser.parseFromString(documentXmlStr, "application/xml");


                // Process comments and document
                let comments = commentsDoc.getElementsByTagName("w:comment");
                // console.log(comments)
                let commentIdsToRemove = [];

                for (let i = 0; i < comments.length; i++) {
                    // console.log(comments[i])
                    if (comments[i].textContent.includes(group)) {
                        const commentId = comments[i].getAttribute("w:id");
                        commentIdsToRemove.push(commentId);
                        comments[i].parentNode.removeChild(comments[i]);
                    }
                }

                console.log("Comment IDs to remove: " + commentIdsToRemove);
                // console.log("Comments to remove: " + commentsToRemoveText);
                console.log("Group: " + group)



                // Use namespace-aware methods to handle elements
                const wNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                let commentStarts = documentDoc.getElementsByTagNameNS(wNamespace, "commentRangeStart");
                let commentEnds = documentDoc.getElementsByTagNameNS(wNamespace, "commentRangeEnd");
                let commentsToRemoveText = [];

                commentIdsToRemove.forEach(commentId => {
                    let start = Array.from(commentStarts).find(elem => elem.getAttribute("w:id") === commentId);
                    let end = Array.from(commentEnds).find(elem => elem.getAttribute("w:id") === commentId);


                    if (start && end) {
                        // Remove all nodes between start and end, including start and end themselves
                        let currentNode = start.nextSibling;
                        // console.log("Removing nodes between start and end: " + start.textContent + " - " + end.textContent)
                        // console.log(currentNode)
                        // console.log(currentNode.textContent)
                        commentsToRemoveText.push(currentNode.textContent);

                        while (currentNode && currentNode !== end) {
                            let nextNode = currentNode.nextSibling;
                            currentNode.parentNode.removeChild(currentNode);
                            currentNode = nextNode;
                        }
                        // Now remove the start and end markers
                        start.parentNode.removeChild(start);
                        end.parentNode.removeChild(end);
                        // -------------------
                        // let parent = commentId.parentNode;
                        // if (parent.nodeName === "w:p" && parent.parentNode) { // Ensure parent's parent exists// Assuming you want to remove the whole paragraph
                        //     console.log("Removing paragraph: " + parent.textContent)
                        //     console.log(parent.parentNode)
                        //     console.log("----------------")
                        //     console.log(parent)
                        //     parent.parentNode.removeChild(parent);
                        // }
                    }
                });
                console.log("Comments to remove: " + commentsToRemoveText);
                console.log(commentsToRemoveText)



                // Use namespace-aware methods to handle elements

                // Serialize both XMLs back to string
                let serializer = new XMLSerializer();
                let updatedCommentsXmlStr = serializer.serializeToString(commentsDoc);
                let updatedDocumentXmlStr = serializer.serializeToString(documentDoc);

                // Replace the old XMLs with the updated ones
                zip.file("word/comments.xml", updatedCommentsXmlStr);
                zip.file("word/document.xml", updatedDocumentXmlStr);
                return zip.generateAsync({ type: "blob" });
            });
        });
}


function modifyDocxComments(blob, group) {
    return JSZip.loadAsync(blob) // Load the .docx file as a zip
        .then(function (zip) {
            const commentsFile = zip.file("word/comments.xml");
            console.log("Comments file: " + commentsFile)
            if (!commentsFile) {
                console.log("No comments file found in the document.");
                // If no comments file, return the original blob unchanged
                return blob;
            }
            return commentsFile.async("string") // Access comments.xml
                .then(function (xmlStr) {
                    // Parse the XML string
                    let parser = new DOMParser();
                    let xmlDoc = parser.parseFromString(xmlStr, "application/xml");
                    console.log(xmlDoc)

                    // Find and remove comments with specific group
                    let comments = xmlDoc.getElementsByTagName("w:comment");
                    console.log
                    console.log(comments)
                    console.log(group)
                    for (let i = 0; i < comments.length; i++) {
                        console.log(comments[i].textContent)
                        // if (comments[i].getAttribute("author") === group) {
                        //     comments[i].parentNode.removeChild(comments[i]);
                        // }
                        if (comments[i].textContent.includes(group)) {
                            console.log("Removing comment: " + comments[i].textContent)
                            comments[i].parentNode.removeChild(comments[i]);
                        }
                    }

                    // Serialize XML back to string
                    let serializer = new XMLSerializer();
                    let updatedXmlStr = serializer.serializeToString(xmlDoc);

                    // Replace the old comments.xml with the updated one
                    zip.file("word/comments.xml", updatedXmlStr);
                    return zip.generateAsync({ type: "blob" });
                });
        });
}

function modifyDocxComments2(blob, group) {
    console.log("Modifying comments for group: " + group)
    console.log(blob)
    return JSZip.loadAsync(blob) // Load the .docx file as a zip
        .then(function (zip) {
            return zip.file("word/comments.xml").async("string"); // Access comments.xml
        })
        .then(function (xmlStr) {
            // Parse the XML string
            console.log("XML String: ")
            console.log(xmlStr)
            console.log("End XML String: ")
            let parser = new DOMParser();
            let xmlDoc = parser.parseFromString(xmlStr, "application/xml");

            // Find and remove comments with specific group
            let comments = xmlDoc.getElementsByTagName("w:comment");
            console.log("Comments: ")
            console.log(comments)
            for (let i = 0; i < comments.length; i++) {
                if (comments[i].getAttribute("author") === group) {
                    comments[i].parentNode.removeChild(comments[i]);
                }
            }

            // Serialize XML back to string
            let serializer = new XMLSerializer();
            let updatedXmlStr = serializer.serializeToString(xmlDoc);

            // Replace the old comments.xml with the updated one
            zip.file("word/comments.xml", updatedXmlStr);
            return zip.generateAsync({ type: "blob" });
        });
}

function removeGroup(html, group) {
    // Regex to remove specific patterns
    const regex = new RegExp(`\\[Group${group}_START\\].*?\\[Group${group}_END\\]`, 'gs');
    return html.replace(regex, '');  // Removes the group content
}

// [GroupA_START] [Group${group}_START\
function saveDocument(content) {
    console.log(content)
    const blob = new Blob([content], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const downloadLink = document.getElementById('downloadLink');
    downloadLink.href = url;
    downloadLink.download = 'modified_document.html';  // Change extension if needed
    downloadLink.style.display = 'inline';
}

function destroyClickedElement(event) {
    document.body.removeChild(event.target);
}

function saveHTM(reportContents) {
    console.log('Generating report...'); // Feedback to user
    try {
        let textToSave = reportNode.replace("##########", reportContents);
        let textToSaveAsBlob = new Blob([textToSave], { type: "text/html" });
        let textToSaveAsURL = URL.createObjectURL(textToSaveAsBlob);
        let fileNameToSaveAs = "AutomateWordReport.htm";

        let downloadLink = document.createElement("a");
        downloadLink.download = fileNameToSaveAs;
        downloadLink.href = textToSaveAsURL;
        downloadLink.style.display = "none";
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink); // Clean up

        console.log('Report generated and download initiated.'); // Success message
    } catch (error) {
        console.error('Error generating report:', error);
        alert('Failed to generate the report. Please try again.'); // Error feedback
    }
}

//saves the current build of the specification hub to a .htm file that can be opened in word and copy pasted in word into a standard Arup Report file
// function saveHTM(reportContents) {
//     var textToSave = reportNode;
//     // var reportContents = document.getElementById("specificationBody").innerHTML;
//     //now add the body of the report into the htm file
//     textToSave = textToSave.replace("##########", reportContents);
//     var textToSaveAsBlob = new Blob([textToSave], { type: "txt/plain" });
//     var textToSaveAsURL = URL.createObjectURL(textToSaveAsBlob);
//     var fileNameToSaveAs = "AutomateWordReport.htm";
//     var downloadLink = document.createElement("a");
//     downloadLink.download = fileNameToSaveAs;
//     downloadLink.innerHTML = "Download File";
//     downloadLink.href = textToSaveAsURL;
//     downloadLink.onclick = destroyClickedElement;
//     downloadLink.style.display = "none";
//     document.body.appendChild(downloadLink);
//     downloadLink.click();
// } 