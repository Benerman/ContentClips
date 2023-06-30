#include "extendscript.csv.jsx"

var console = console || {};
console.log = function (message) {
    $.writeln(message);
}

getSep = function () {
    if (Folder.fs === 'Macintosh') {
        return '/';
    } else {
        return '\\';
    }
}


var _sep = ',';
var _newline = '\n';



// var CSV = function() {};

// CSV.openFile = function(filePath){
//     csvFile = File(filePath)
//     if (csvFile !== null) {
//         csvFile.open('r');
//         var content = csvFile.read();
//         csvFile.close();
//         return content;
//     }
// };

// CSV.splitContents = function (contents) { //_sep=',', _newline='\n') {
//     // Create CSV PARSING
//     var rows = contents.split(_newline);
//     console.log(rows.length)
//     var splitRows = {};
//     splitRows.rows = []
//     splitRows.header = []
//     for (i=0;i<rows.length;i++) {
//         if (i === 0) {
//             header = rows[0].split(_sep);
//             splitRows.header = header
//             continue
//         } else {
//             var row = {}
//             row.cells = []
//             var cells = rows[i].split(_sep)
//             for (n=0;n<cells.length; n++) {
//                 console.log(cells[n])
//                 row.cells[n] = cells[n]
//             }
//             splitRows.rows[i-1] = row
//             console.log(row)
//         }
//     }
//     console.log(splitRows.toString())
//     return splitRows
// };

var DEBUG = false

console.log('Open File')
var useDialog = DEBUG !== true // Make it so useDialog is True when DEBUG is FALSE
var csvContents = CSV.toJSON("/Users/blarson/Desktop/four slide template/Four\ Slide\ Data.csv", useDialog, separator=_sep)



var mainSequenceName = "Sequence 01"

var encodePresetPath = "/Users/blarson/Desktop/four slide template/CC Export Preset.epr";




var graphicComponentNum = 2;
var graphicPropertiesEnum = {
    0: "Background Color",
    1: "Image",
    2: "Text",
    3: "Font controls",
    4: "Text Spacing",
    5: "Text Position",
    6: "Spreadsheet"
};
var imageNameEnum = {
    "Image 1.jpg": "Product Main Image",
    "Image 2.jpg": "Slide 1 Image",
    "Image 3.jpg": "Slide 2 Image",
    "Image 4.jpg": "Slide 3 Image",
};

var imageNamesToChange = ["Image 1.jpg", "Image 2.jpg", "Image 3.jpg", "Image 4.jpg"]
// vTracks[0].clips[1].components[2].properties[0].getColorValue()
// for (i=0;i<vTracks[0].clips[1].components[2].properties[1].numItems;i++) {vTracks[0].clips[1].components[2].properties[1].getValue()}

// app.project.rootItem.children[9].canChangeMediaPath("/Users/blarson/Desktop/four slide template/On Deck Images")



if (DEBUG === true) {
    var projectFile = "/Users/blarson/Desktop/four slide template/four slide template.prproj"
} else {
    var projectFile = File.openDialog("Select a Premiere Project file to import.", "Premiere Project File:*.prproj;All files:*.*", false);
}
console.log(projectFile.fsName)
console.log(projectFile.fullName)
console.log(projectFile.absoluteURI)
console.log(projectFile.path)
app.openDocument(projectFile.absoluteURI)

var mainSequence = app.project.activeSequence;

app.encoder.launchEncoder()
// $.sleep(5)

for (x=0;x<csvContents.length;x++) {
    // Need to address the text for the Mogrt
    // do it here
    console.log(mainSequence.videoTracks.length)
    for (i=0;i<mainSequence.videoTracks.length;i++) {
        var curTrack = mainSequence.videoTracks[i]
        for (o=0;o<curTrack.clips.length;o++) {
            console.log("Clip: "+o)
            // Video Track Loops
            // Video Track 1
            if (i === 0) {
                // slide 2, 4, 6
                // ignore first slide(slide 2)
                if (o === 1) { // || o ===2) {
                    // Update Slide 2's text parameter
                    if (o === 1) {
                        curTrack.clips[o].components[2].properties[3].setValue(csvContents[x]["Slide 2"].toString(), 1)
                    }
                    // else {
                    //     // for future use 
                    //     curTrack.clips[o].components[2].properties[2].setValue(csvContents[x]["Slide 4"])
                    // }
                }
            // Video Track 1
            } else if (i === 1) {
                // slide 1, 3, 5
                // Update Slide 2's text parameter
                if (o === 0) {
                    var text = csvContents[x]["Product Name"]
                    curTrack.clips[o].components[2].properties[0].setValue(csvContents[x]["Product Name"].toString(), 1)
                    curTrack.clips[o].components[2].properties[1].setValue(csvContents[x]["Product Description"].toString(), 1)
                } else if (o === 1) {
                    curTrack.clips[o].components[2].properties[3].setValue(csvContents[x]["Slide 1"].toString(), 1)
                } else if (o === 2) {
                    curTrack.clips[o].components[2].properties[3].setValue(csvContents[x]["Slide 3"].toString(), 1)
                }
            }
        }
    }
    
    // This addresses the Image Swap
    // Update Images in the MOGRT, They are Sequences name "Image"

    var allSequences = app.project.sequences
    for (i=0;i<allSequences.length;i++) {
        if (allSequences[i].name === "Image") {
            // Update source image with one provided
            if (allSequences[i].videoTracks.length === 1 && allSequences[i].videoTracks[0].clips.length === 1) {
                var imgName = allSequences[i].videoTracks[0].clips[0].name
                // allSequences[i].videoTracks[0].clips[0].projectItem.changeMediaPath("/Users/blarson/Desktop/four slide template/On Deck Images/" + imgName.toString() )
                allSequences[i].videoTracks[0].clips[0].projectItem.changeMediaPath(csvContents[x][imageNameEnum[imgName]])
            }
        }
    }
    // Add to Encoder
    app.encoder.encodeSequence(mainSequence, "/Users/blarson/Desktop/" + csvContents[x]["Product Name"]+".mp4", encodePresetPath, workArea=0)
    // app.project.saveAs("./" + csvContents[x]['Product Name'].replace(" ", "_") + ".prproj")
    console.log("Done")
}




// if (app.project.rootItem.children[9].canChangeMediaPath() === 0) {
//     console.log(app.project.rootItem.children[9].canChangeMediaPath("/Users/blarson/Desktop/four slide template/On Deck Images") === 0)
// }

// vTracks[0].clips[1].components
// console.log(firstVTrack.name)








// function cloneAndRename(app, sequenceToClone, newName) {
//     // Clone Main sequence
//     console.log(sequenceToClone.clone());

//     var allSequences = app.project.sequences;
//     for (i = 0; i < allSequences.length; i++) {
//         var seq = app.project.sequences[i]
//         if (seq.name == mainSequenceName+" Copy") {
//             if (seq.name != "Image") {
//                 if (seq.name.indexOf('Copy') != -1) {
//                     seq.name = newName
//                     newSequence = seq
//                 }
//             }
//         } else {
//             console.log("Skipping Sequence, Main Sequence")
//             continue
//         }
//     }
//     app.project.activeSequence = newSequence
//     return newSequence
// }

// var newSequence = cloneAndRename(app, mainSequence, "Test Sequence")

// // console.log(mainSequence.clone());
// // var allSequences = app.project.sequences;

// // for (i = 0; i < allSequences.length; i++) {
// //     var seq = app.project.sequences[i]
// //     if (seq.name == mainSequenceName) {
// //         console.log("Skipping Sequence, Main Sequence")
// //         continue
// //     } else {
// //         if (seq.name != "Image" && seq.name.indexOf('Copy') != -1) {
// //             seq.name = sequenceNames[seqNameNumber]
// //             seqNameNumber ++;
// //         }
// //     }
// // }


// // Evaluate tracks within Sequence
// console.log(app.project.activeSequence.name)
// console.log(app.project.rootItem.clips[0].name)



// // for (seqName of sequenceNames) {
// //     var seq = app.project.activeSequence
// //     seq.name = seqName
// // }

// // var clips = app.project.rootItem.children


// // videoTracks = app.project.activeSequence.videoTracks;
// // // console.log(videoTracks)
// // firstTrack = videoTracks[0];
// // firstTrack.setMute(true);
