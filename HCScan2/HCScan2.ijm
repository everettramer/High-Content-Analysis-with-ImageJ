//Global variables
//Assay parameters
var nFieldRows;          //int 1 - 8
var nFieldCols;           //int 1 - 8
var analyzeField;        //int 0 - 64
//Channels 1 & 2
var rollingBallRadius; //pixels
var thresholdMethod; //string
var thresholdAdjust;   //% (or lower level if manual threshold method)
//Channel 1
var smoothRadius;    //pixels
var minArea; //pixels
var maxArea;             //pixels
var minCircularity;     //float 0.0 - 1.0
//Channel 2
var ringWidth;            //pixels
//Flags
var loadDefaults;                          //Boolean
var saveDefaults;                         //Boolean
var stopAfterLoadingImages;     //Boolean
var stopAfterValidObjects;          //Boolean
var stopAfterLabeledMasks;      //Boolean
var stopForAnalysisReview;       //Boolean
//var valid;			       //Boolean

//Other globals
//var historyPath = getDirectory("temp") + "\\HCScan2\\HCScan2 History.txt";
var defaultsPath = getDirectory("home") + "\\Documents\\HCScan2\\HCScan2 Default Values.txt";
var imagePath;
var imageTitle;
var imageType;
var analysisDateStamp;
var analysisTimeStamp;
var reportDirectory;
var reportXLS;
var reportFile;
var openReportFile;                     //Boolean
var currentField;
var fieldX1, fieldY1, fieldWidth, fieldHeight; //Field Upper left corner (X1, Y1) and size - units are pixels
var thresholdLower, thresholdUpper;

macro "HCScan2..." {
  openReportFile = true;
  //Get assay parameter values---------------------------------------------------------------------------------------------------
  do {
    loadDefaultValues();
    showAssayParamsDialog();
    if (loadDefaults) defaultsPath = File.openDialog("Open Defaults File");
    else if (saveDefaults) saveDefaultValues();
  } while ( loadDefaults || saveDefaults );
  minArea = maxOf(minArea, 2); //Always have at least 2 circ pixels
  minRingArea = maxOf(floor(3 * (sqrt(minArea) + ringWidth * ringWidth)), 2); //Always have at least 2 ring pixels
  cleanupAfterAnalysis = !(stopAfterLoadingImages || stopAfterValidObjects || stopAfterLabeledMasks);

  //Get image path, title, and type
  close("*"); //Close all images
  run("Open..."); //Propmt user to open either CH1 or CH2 of image
  run("Set Scale...", "distance=0 known=0 pixel=1 unit=pixel global"); //Set all units to pixels
  //Get analysis date & time stamp
  analysisDateStamp = getDateStamp();
  analysisTimeStamp = getTimeStamp();
  //Get path
  imagePath = File.directory;
  imageTitleAndType = getTitle();
  //Get title and type
  index = lastIndexOf(imageTitleAndType, ".");
  if (index > 4) {
    imageTitle = substring(imageTitleAndType, 0, index - 4);
    imageType = substring(imageTitleAndType, index + 1, lengthOf(imageTitleAndType));
    ch1Title = imageTitle + "_CH1";
    ch2Title = imageTitle + "_CH2";

    //Open image pair
    close("*"); //Close all images
    open(imagePath + "//" + ch1Title + "." + imageType);
    rename(ch1Title);
    open(imagePath + "//" + ch2Title + "." + imageType);
    rename(ch2Title);

    currentField = 1;
    if (nFieldRows == 1 && nFieldCols == 1) {
      fieldX1 = 0;
      fieldY1 = 0;
      fieldWidth = getWidth();
      fieldHeight = getHeight();
      analyze(ch1Title, ch2Title);
    }
    else {
      stopAfterFirstField = stopAfterLoadingImages || stopAfterValidObjects || stopAfterLabeledMasks;
      stopAfterFirstField = stopAfterFirstField || stopForAnalysisReview || analyzeField != 0;
      for (r = 0; r < nFieldRows; r++) {
        for (c = 0; c < nFieldCols; c++) {
          if (analyzeField == 0 || analyzeField == currentField) {
            analyze(cropImage(ch1Title, c, r), cropImage(ch2Title, c, r));
            if (stopAfterFirstField) {
              r = nFieldRows + 1; //This will stop the row & column loops
              c = nFieldCols + 1;
            }
          }
          currentField++;
        }
      }
    }
  }
  else print("Invalid image file name");

  if (!openReportFile) File.close(reportFile);
  if (cleanupAfterAnalysis) cleanup("Analysis Complete", ch1Title, ch2Title);
}

function analyze(ch1Title, ch2Title) {
  //  valid = true;
  if (stopAfterLoadingImages) {
    return;
  }

  //Remove background in both channels
  selectWindow(ch1Title);
  run("Subtract Background...", "rolling=" + rollingBallRadius + " disable"); //Disable smoothing
  selectWindow(ch2Title);
  run("Subtract Background...", "rolling=" + rollingBallRadius + " disable"); //Disable smoothing
  //Smooth
  selectWindow(ch1Title);
  run("Duplicate...", "title=Circle");
  //Smooth to aggregate objects and smooth their boundaries
  run("Gaussian Blur...", "sigma=" + smoothRadius);

  //Segment objects from background
  threshold("Circle", thresholdMethod, thresholdAdjust);

  //Create binary masks of valid objects-----------------------------------------------------------------------------------------------------
  //Circles
  maxArea = minOf(maxArea, getWidth() / 2.0 * getHeight() / 2.0);
  roiManager("reset");
  run("Clear Results");
  run("Set Measurements...", " redirect=None decimal=2");
  run("Analyze Particles...", "size=" + minArea + "-" + maxArea + " pixel circularity=" + minCircularity + "-1.00 show=Masks clear add exclude include in_situ");
  nValidObjects = roiManager("count");
  print("nValidObjects " + nValidObjects);

  if (stopAfterValidObjects) {
    return;
  }

  if (nValidObjects > 0) {
    //Voronoi - Used to separate adjacent circles after their dilation
    selectWindow("Circle");
    run("Duplicate...", "title=Voronoi");
    selectWindow("Voronoi");
    run("Options...", "count=1 edm=Overwrite do=Nothing");
    run("Voronoi");
    setThreshold(1, 255);
    run("Convert to Mask");
    //run("Options...", "iterations=1 count=1 edm=Overwrite do=Nothing");
    //run("Dilate");
    run("Invert");

    //Rings
    selectWindow("Circle");
    run("Duplicate...", "title=Ring");
    //Dilate the circles
    run("Maximum...", "radius=" + ringWidth);
    //Separate any that might have touched during dilation
    imageCalculator("AND", "Ring", "Voronoi");

    //Create labeled Voronoi--------------------------------------------------------------------------------------------------------------
    selectWindow("Voronoi");
    //Set min Voronoi size to minArea pixels to ignore debris created during circ dilation and Voronoi creation
    roiManager("reset");
    run("Clear Results");
    run("Set Measurements...", " redirect=None decimal=2");
    run("Analyze Particles...", "size=" + minArea + "-Infinity circularity=0.00-1.00 show=Nothing display clear add include in_situ");
    //run("Analyze Particles...", "size=10-Infinity circularity=0.00-1.00 show=Nothing display clear add include in_situ");
    nVoronoi = roiManager("count");
    print("nVoronoi " + nVoronoi);
    if (nVoronoi != nValidObjects) {
      //     valid = false;
      message = "Field " + currentField + ": nVoronoi (" + nVoronoi + ") does not equal nValidObjects (" + nValidObjects + "): Analysis aborted";
      if (cleanupAfterAnalysis) {
        cleanup(message, ch1Title, ch2Title);
      }
      return;
    }
    //Fill Voronoi regions with a gray scale value equal to their ROI index
    run("16-bit");
    for (i = 0; i < nValidObjects; i++) {
      roiManager("Select", i);
      setColor(i + 1); //Add 1 because index starts at 0, but 0 is background value
      fill();
    }
    run("Select None");

    //Create circle ROI's with Voronoi labels--------------------------------------------------------------------------------------------------------------
    selectWindow("Circle");
    roiManager("reset");
    run("Clear Results");
    run("Set Measurements...", "modal redirect='Voronoi' decimal=0");
    run("Analyze Particles...", "size=0-Infinity circularity=0.00-1.00 show=Nothing clear add include");
    //combineROIFragments(0, nVoronoi);
    //Change circle ROI labels to Voronoi labels
    nCirc = roiManager("count");
    print("nCirc " + nCirc);
    if (nCirc != nVoronoi) {
      //		valid = false;
      message = "(Field " + currentField + "): nCirc (" + nCirc + ") does not equal nVoronoi (" + nVoronoi + "): Analysis aborted";
      print(message);
      if (cleanupAfterAnalysis) {
        cleanup(message, ch1Title, ch2Title);
      }
      return;
    }
    circIndices = newArray(nCirc); //Indices of circ masks in ROI Manager
    renameROIs('C', 0, nCirc, circIndices);

    //Create circle ROI's with Voronoi labels--------------------------------------------------------------------------------------------------------------
    selectWindow("Ring");
    run("Set Measurements...", "modal redirect='Voronoi' decimal=0");
    run("Analyze Particles...", "size=0-Infinity circularity=0.00-1.00 show=Nothing add include");
    //combineROIFragments(nCirc, nVoronoi);
    //Change ring mask labels to Voronoi labels
    nRing = roiManager("count") - nCirc; //Ring indices run nCirc..nRing + nCirc
    print("nRing " + nRing);
    if (nRing != nVoronoi) {
      //		valid = false;
      message = "(Field " + currentField + "): nRing (" + nRing + ") does not equal nVoronoi (" + nVoronoi + "): Analysis aborted";
      if (cleanupAfterAnalysis) {
        cleanup(message, ch1Title, ch2Title);
      }
      return;
    }
    ringIndices = newArray(nRing); //Indices of ring masks in ROI Manager
    renameROIs('D', nCirc, nRing, ringIndices); //Use D because it follows C in alphabetical sort
    //The temporary 'D' ROI's have only outer boundaries
    //XOR with circle ROI's to make true ring (with inner & outer boundaries)
    selectedROIs = newArray(2);
    for (i = 0; i < nRing; i++) {
      selectedROIs[0] = i; //circle
      selectedROIs[1] = nCirc + i; //ring	//Ring indices run nCirc..nRing + nCirc
      roiManager("Select", selectedROIs);
      roiManager("XOR");
      roiManager("Add", "ffff00");
      iCombined = nCirc + nRing + i;
      roiManager("Select", iCombined);
      roiManager("Rename", leftPad("R", i + 1, nRing));
    }
    //Delete temporary 'D' rings
    for (i = nCirc; i < nCirc + nRing; i++) { //Ring indices run nCirc..nRing + nCirc
      roiManager("Select", nCirc); //Use nCirc because ROI manager list shifts down as entries are deleted
      roiManager("Delete");
    }

    if (stopAfterLabeledMasks) {
      return;
    }

    //Validate ring and circ labels before performing the analysis
    run("Set Measurements...", "modal redirect='Voronoi' decimal=0");
    run("Clear Results");
    roiManager("select", circIndices);
    roiManager("Measure");
    roiManager("select", ringIndices);
    roiManager("Measure");
    for (i = 0; i < nVoronoi; i++) {
      circLabel = getResult("Mode", i);
      ringLabel = getResult("Mode", nCirc + i); //Ring indices run nCirc..nRing + nCirc
      if (circLabel != ringLabel) {
        //		  valid = false;
        message = "Field " + currentField + ": circ and ring labels for object " + i + " do not agree: Analysis aborted";
        if (cleanupAfterAnalysis) cleanup(message, ch1Title, ch2Title);
        return;
      }
    }

    //Analyze----------------------------------------------------------------------------------------------------------------------
    //Circles - Channel 1
    selectWindow(ch1Title);
    //run("Set Measurements...", "area feret's shape mean standard min max integrated add redirect='"+ch1Title+"' decimal=2");
    run("Set Measurements...", "area feret's shape mean standard min max integrated add redirect=None decimal=2");
    run("Clear Results");
    roiManager("select", circIndices);
    roiManager("Measure");
    run("Remove Overlay");
    //saveAs("Measurements", CircleXLS);
    circArea = newArray(nCirc);
    circFeretDiameter = newArray(nCirc);
    circCircularity = newArray(nCirc);
    ch1CircMean = newArray(nCirc);
    ch1CircStdDev = newArray(nCirc);
    ch1CircMin = newArray(nCirc);
    ch1CircMax = newArray(nCirc);
    ch1CircSumPixels = newArray(nCirc);
    for (i = 0; i < nCirc; i++) {
      circArea[i] = getResult("Area", i);
      circCircularity[i] = getResult("Circ.", i);
      circFeretDiameter[i] = getResult("Feret", i);
      ch1CircMean[i] = getResult("Mean", i);
      ch1CircStdDev[i] = getResult("StdDev", i);
      ch1CircMin[i] = getResult("Min", i);
      ch1CircMax[i] = getResult("Max", i);
      ch1CircSumPixels[i] = getResult("RawIntDen", i);
    }

    //Circles - Channel 2
    selectWindow(ch2Title);
    run("Set Measurements...", "mean standard min max integrated redirect=None decimal=2");
    run("Clear Results");
    roiManager("select", circIndices);
    roiManager("Measure");
    run("Remove Overlay");
    //saveAs("Measurements", CircleXLS);
    ch2CircMean = newArray(nCirc);
    ch2CircStdDev = newArray(nCirc);
    ch2CircMin = newArray(nCirc);
    ch2CircMax = newArray(nCirc);
    ch2CircSumPixels = newArray(nCirc);
    for (i = 0; i < nCirc; i++) {
      ch2CircMean[i] = getResult("Mean", i);
      ch2CircStdDev[i] = getResult("StdDev", i);
      ch2CircMin[i] = getResult("Min", i);
      ch2CircMax[i] = getResult("Max", i);
      ch2CircSumPixels[i] = getResult("RawIntDen", i);
    }

    if (cleanupAfterAnalysis) {
      selectWindow("Circle");
      close();
    }

    //Rings - Channel 2
    selectWindow(ch2Title);
    //run("Set Measurements...", "area mean standard min max integrated add redirect='"+ch2Title+"' decimal=2");
    run("Set Measurements...", "area mean standard min max integrated add redirect=None decimal=2");
    run("Clear Results");
    roiManager("select", ringIndices);
    roiManager("Measure");
    run("Remove Overlay");
    //saveAs("Measurements", ch2RingXLS);
    ringArea = newArray(nRing);
    ch2RingMean = newArray(nRing);
    ch2RingStdDev = newArray(nRing);
    ch2RingMin = newArray(nRing);
    ch2RingMax = newArray(nRing);
    ch2RingSumPixels = newArray(nRing);
    for (i = 0; i < nRing; i++) {
      ringArea[i] = getResult("Area", i);
      ch2RingMean[i] = getResult("Mean", i);
      ch2RingStdDev[i] = getResult("StdDev", i);
      ch2RingMin[i] = getResult("Min", i);
      ch2RingMax[i] = getResult("Max", i);
      ch2RingSumPixels[i] = getResult("RawIntDen", i);
    }

    if (cleanupAfterAnalysis) {
      selectWindow("Ring");
      close();
      selectWindow("Voronoi");
      close();
    }

    //Display overlays on ch1 & ch2 images------------------------------------------------------------------------------------------
    roiManager("UseNames", "true");
    selectWindow(ch1Title);
    run("Overlay Options...", "stroke=red width=2 fill=none apply");
    roiManager("Show All with labels");
    selectWindow(ch2Title);
    run("Overlay Options...", "stroke=green width=2 fill=none apply");
    roiManager("Show All with labels");
  }
  else {
    if (cleanupAfterAnalysis) {
      selectWindow("Circle");
      close();
    }
  }

  //Create Results table----------------------------------------------------------------------------------------------------------------------
  run("Clear Results");
  for (i = 0; i < nValidObjects; i++) {
    //setResult("Label",i,circLabel[i]);
    setResult("Circ Area", i, circArea[i]);
    setResult("Circ Circularity", i, circCircularity[i]);
    setResult("Ring Area", i, ringArea[i]);
    setResult("ch1 Circ Mean", i, ch1CircMean[i]);
    setResult("ch1 Circ StdDev", i, ch1CircStdDev[i]);
    setResult("ch1 Circ Sum", i, ch1CircSumPixels[i]);

    //setResult("Ring Label",i,ringLabel[i]);
    setResult("ch2 Circ Mean", i, ch2CircMean[i]);
    setResult("ch2 Circ StdDev", i, ch2CircStdDev[i]);
    setResult("ch2 Circ Sum", i, ch2CircSumPixels[i]);
    setResult("ch2 Ring Mean", i, ch2RingMean[i]);
    setResult("ch2 Ring StdDev", i, ch2RingStdDev[i]);
    setResult("ch2 Ring Sum", i, ch2RingSumPixels[i]);
  }
  updateResults();

  //Create report----------------------------------------------------------------------------------------------------------------------
  if (openReportFile) {
    openReportFile = false;
    reportDirectory = imagePath + imageTitle + "_Report" + "_" + analysisDateStamp + "_" + analysisTimeStamp;
    File.makeDirectory(reportDirectory)
    reportXLS = reportDirectory +  "\\" + imageTitle + ".xls";
    reportFile = File.open(reportXLS);
    print(reportFile, "Image:\t" + imagePath + imageTitle + "." + imageType);
    print(reportFile, "Field:\t" + currentField + "\n");
    print(reportFile, "Analysis Date:\t" + analysisDateStamp + "\n");
    print(reportFile, "Analysis Time:\t" + analysisTimeStamp + "\n");
    print(reportFile, "\n");

    print(reportFile, "Analysis Parameters:\n");
    print(reportFile, "Field Rows\t" + nFieldRows + "\n");
    print(reportFile, "Field Cols\t" + nFieldCols + "\n");
    print(reportFile, "Rolling Ball Radius (pixels)\t" + rollingBallRadius + "\n");
    print(reportFile, "Threshold Method\t" + thresholdMethod + "\n");
    print(reportFile, "Threshold Adjust (%)\t" + thresholdAdjust + "\n");
    print(reportFile, "Threshold Lower\t" + thresholdLower + "\n");
    print(reportFile, "Threshold Upper\t" + thresholdUpper + "\n");
    print(reportFile, "Smooth Radius (pixels)\t" + smoothRadius + "\n");
    print(reportFile, "Min Area (pixels)\t" + minArea + "\n");
    print(reportFile, "Max Area (pixels)\t" + maxArea + "\n");
    print(reportFile, "Min Circularity\t" + minCircularity + "\n");
    print(reportFile, "Ring Width (pixels)\t" + ringWidth + "\n");
    print(reportFile, "\n");
  }

  print(reportFile, "Analysis Field:\t" + currentField + "\n");
  print(reportFile, "Upper Left Corner\t (" + fieldX1 + "," + fieldY1 + ")\n");
  print(reportFile, "Width\t" + fieldWidth + "\n");
  print(reportFile, "Height\t" + fieldHeight + "\n");
  print(reportFile, "\n");

  print(reportFile, "Analysis Results:\n");
  headerLine = "";
  headerLine = headerLine + "ROI\t";
  headerLine = headerLine + "Circ Area\t";
  headerLine = headerLine + "Circ Feret Diameter\t";
  headerLine = headerLine + "Circ Circularity\t";
  headerLine = headerLine + "Ring Area\t";
  headerLine = headerLine + "ch1 Circ Mean\t";
  headerLine = headerLine + "ch1 Circ StdDev\t";
  headerLine = headerLine + "ch1 Circ Min\t";
  headerLine = headerLine + "ch1 Circ Max\t";
  headerLine = headerLine + "ch1 Circ Sum Pixels\t";
  headerLine = headerLine + "ch2 Circ Mean\t";
  headerLine = headerLine + "ch2 Circ StdDev\t";
  headerLine = headerLine + "ch2 Circ Min\t";
  headerLine = headerLine + "ch2 Circ Max\t";
  headerLine = headerLine + "ch2 Circ Sum Pixels\t";
  headerLine = headerLine + "ch2 Ring Mean\t";
  headerLine = headerLine + "ch2 Ring StdDev\t";
  headerLine = headerLine + "ch2 Ring Min\t";
  headerLine = headerLine + "ch2 Ring Max\t";
  headerLine = headerLine + "ch2 Ring Sum Pixels\n";
  print(reportFile, headerLine);

  for (i = 0; i < nValidObjects; i++) {
    line = "";
    line = line + (i + 1) + "\t";
    line = line + circArea[i] + "\t";
    line = line + circFeretDiameter[i] + "\t";
    line = line + circCircularity[i] + "\t";
    line = line + ringArea[i] + "\t";
    line = line + ch1CircMean[i] + "\t";
    line = line + ch1CircStdDev[i] + "\t";
    line = line + ch1CircMin[i] + "\t";
    line = line + ch1CircMax[i] + "\t";
    line = line + ch1CircSumPixels[i] + "\t";
    line = line + ch2CircMean[i] + "\t";
    line = line + ch2CircStdDev[i] + "\t";
    line = line + ch2CircMin[i] + "\t";
    line = line + ch2CircMax[i] + "\t";
    line = line + ch2CircSumPixels[i] + "\t";
    line = line + ch2RingMean[i] + "\t";
    line = line + ch2RingStdDev[i] + "\t";
    line = line + ch2RingMin[i] + "\t";
    line = line + ch2RingMax[i] + "\t";
    line = line + ch2RingSumPixels[i] + "\t";
    print(reportFile, line);
  }
  print(reportFile, "\n");

  //Save images with ROI's
  selectWindow(ch1Title);
  ch1Image = reportDirectory +  "\\" + imageTitle + "_CH1_F" + currentField;
  saveAs("zip", ch1Image);
  rename(ch1Title);
  selectWindow(ch2Title);
  ch2Image = reportDirectory +  "\\" + imageTitle + "_CH2_F" + currentField;
  saveAs("zip", ch2Image);
  rename(ch2Title);

  if (cleanupAfterAnalysis) {
    if (!stopForAnalysisReview) {
      if (isOpen(ch1Title)) {
        selectWindow(ch1Title);
        close();
      }
      if (isOpen(ch2Title)) {
        selectWindow(ch2Title);
        close();
      }
    }
  }
}
}

/*
function loadHistory() {
  // Get path to temp directory
  hp = getDirectory(historyPath);
  if (hp=="") {
    // Create a history directory in temp
    File.makeDirectory(historyPath);
    if (!File.exists(historyPath + historyName))
      exit("Unable to create history directory");
  }
  if (File.exists(historyName))
    historyString = File.openAsString(defaultsPath + defaultsName);
    historyLines = split(defaultsString,"\n");
    defaultsPath = defaultsLines[0];
    defaultsName = defaultsLines[1];
    imagePath = defaultsLines[2];
  }
  
function saveHistory(fullName) {
  historyString = "";
  historyString = historyString + defaultsPath + "\n";
  historyString = historyString + defaultsName + "\n";
  historyString = historyString + imagePath + "\n";
  File.saveString(historyString, fullName);
  }
*/
}

//Helper Functions-------------------------------------------------------------------------------------------------------------------------------
function loadDefaultValues() {
  defaultsString = File.openAsString(defaultsPath);
  defaultsLines = split(defaultsString, "\n");
  nFieldRows = parseInt(defaultsLines[0]);
  nFieldCols = parseInt(defaultsLines[1]);
  analyzeField = parseInt(defaultsLines[2]);
  rollingBallRadius = parseInt(defaultsLines[3]);
  smoothRadius = parseInt(defaultsLines[4]);
  thresholdMethod = defaultsLines[5];
  thresholdAdjust = parseFloat(defaultsLines[6]);
  minArea = parseInt(defaultsLines[7]);
  maxArea = parseInt(defaultsLines[8]);
  minCircularity = parseFloat(defaultsLines[9]);
  ringWidth = parseInt(defaultsLines[10]);
}

function saveDefaultValues() {
  defaultsString = "";
  defaultsString = defaultsString + nFieldRows + " nFieldRows\n";
  defaultsString = defaultsString + nFieldCols + " nFieldCols\n";
  defaultsString = defaultsString + analyzeField + " analyzeField\n";
  defaultsString = defaultsString + rollingBallRadius + " rollingBallRadius\n";
  defaultsString = defaultsString + smoothRadius + " smoothRadius\n";
  defaultsString = defaultsString + thresholdMethod + "\n";
  defaultsString = defaultsString + thresholdAdjust + " thresholdAdjust\n";
  defaultsString = defaultsString + minArea + " minArea\n";
  defaultsString = defaultsString + maxArea + " maxArea\n";
  defaultsString = defaultsString + minCircularity + " minCircularity\n";
  defaultsString = defaultsString + ringWidth + " ringWidth\n";
  defaultsPath = File.openDialog("Save Defaults File");
  File.saveString(defaultsString, defaultsPath);
}

function showAssayParamsDialog() {
  Dialog.create("Assay Parameters");
  Dialog.addNumber("# of Field Rows:", nFieldRows, 0, 8, "images");
  Dialog.addNumber("# of Field Cols:", nFieldCols, 0, 8, "images");
  Dialog.addNumber("Analyze Field:", analyzeField, 0, 8, "(0 = All Fields)");
  Dialog.addCheckbox("Stop after Loading Images", false);
  Dialog.addCheckbox("Stop after Valid Objects", false);
  Dialog.addCheckbox("Stop after Labeled Masks", false);
  Dialog.addCheckbox("Stop for Analysis Review", false);
  Dialog.addNumber("Rolling Ball Radius:", rollingBallRadius, 0, 8, "pixels");
  Dialog.addNumber("Smoothing Radius:", smoothRadius, 0, 8, "pixels");
  Dialog.addChoice("Threshold Method:", newArray("MaxEntropy", "Triangle", "IsoData", "Manual"), thresholdMethod);
  Dialog.addNumber("Threshold Adjust:", thresholdAdjust, 0, 8, "%/level");
  Dialog.addNumber("Min Area:", minArea, 0, 8, "pixels");
  Dialog.addNumber("Max Area:", maxArea, 0, 8, "pixels");
  Dialog.addNumber("Min Circularity:", minCircularity, 2, 8, "");
  Dialog.addNumber("Ring Width:", ringWidth, 0, 8, "pixels");
  Dialog.addCheckbox("Load Default Values", false);
  Dialog.addCheckbox("Save Default Values", false);
  Dialog.show();
  nFieldRows = Dialog.getNumber();
  nFieldCols = Dialog.getNumber();
  analyzeField = Dialog.getNumber();
  stopAfterLoadingImages = Dialog.getCheckbox();
  stopAfterValidObjects = Dialog.getCheckbox();
  stopAfterLabeledMasks = Dialog.getCheckbox();
  stopForAnalysisReview = Dialog.getCheckbox();
  rollingBallRadius = Dialog.getNumber();
  smoothRadius = Dialog.getNumber();
  thresholdMethod = Dialog.getChoice();
  thresholdAdjust = Dialog.getNumber();
  minArea = Dialog.getNumber();
  maxArea = Dialog.getNumber();
  minCircularity = Dialog.getNumber(); //Enforce this value to be between 0.0 and 1.0
  ringWidth = Dialog.getNumber();
  loadDefaults = Dialog.getCheckbox();
  saveDefaults = Dialog.getCheckbox();
}

function cropImage(imageTitle, c, r) {
  colrowString = "" + leftPad("_R", r, nFieldRows) + leftPad("_C", c, nFieldCols);
  selectWindow(imageTitle);
  fieldTitle = imageTitle + colrowString;
  run("Duplicate...", "'title=[" + fieldTitle + "]'");
  h = getHeight();
  w = getWidth();
  fieldX1 = abs(floor(c * w / nFieldCols));
  fieldY1 = abs(floor(r * h / nFieldRows));
  //print(x1); print(y1); 
  fieldWidth = abs(floor(w / nFieldCols));
  fieldHeight = abs(floor(h / nFieldRows));
  //print(width);print(height); 
  makeRectangle(fieldX1, fieldY1, fieldWidth, fieldHeight);
  run("Crop");
  return fieldTitle;
}

function threshold(chTitle, method, adjust) {
  selectWindow(chTitle);
  if (method == "Manual") {
    setThreshold(adjust, 255);
  }
  else {
    setAutoThreshold(method + " dark");
    getThreshold(lower, upper);
    lower = lower * (1.0 + adjust / 100.0);
    setThreshold(lower, 255);
  }
  getThreshold(thresholdLower, thresholdUpper);
  run("Convert to Mask");
}

function leftPad(label, value, maxValue) { //Formats numbers with leading zeros
  maxValueString = "" + maxValue;
  valueString = "" + value;
  for (i = lengthOf(valueString); i < lengthOf(maxValueString); i++) {
    label = label + 0;
  }
  s = label + value;
  return s;
}

function getDateStamp() {
  getDateAndTime(year, month, dayOfWeek, dayOfMonth, hour, minute, second, msec);
  dateString = "Y" + year + "M";
  month++;
  if (month < 10) {
    dateString = dateString + "0";
  }
  dateString = dateString + month + "D";
  if (dayOfMonth < 10) {
    dateString = dateString + "0";
  }
  dateString = dateString + dayOfMonth;
  return dateString;
}

function getTimeStamp() {
  getDateAndTime(year, month, dayOfWeek, dayOfMonth, hour, minute, second, msec);
  timeString = "h";
  if (hour < 10) {
    timeString = timeString + "0";
  }
  timeString = timeString + hour + "m";
  if (minute < 10) {
    timeString = timeString + "0";
  }
  timeString = timeString + minute + "s";
  if (second < 10) {
    timeString = timeString + "0";
  }
  timeString = timeString + second;
  return timeString;
}

function combineROIFragments(offset, nVoronoiLabels) {
  //offset skips the first part of the ROI manager list
  nROIFragments = roiManager("count");
  //Combine ROI fragments that have identical Voronoi labels
  selectedROIs = newArray(2);
  //Create a list of combined ROI indices, one for each label
  combinedROIIndices = newArray(nVoronoiLabels);
  Array.fill(combinedROIIndices, -1); //-1 => no ROI fragment assigned to this label
  //Scan list of ROI fragments and assign to label
  for (i = offset; i < nROIFragments; i++) {
    roiLabel = getResult("Mode", i);
    if (combinedROIIndices[roiLabel - 1] == -1) {
      //No combined ROI started for this label => assign current fragment
      combinedROIIndices[roiLabel - 1] = i;
    }
    else {
      //Combined ROI already started for this label=> combine current ROI fragment with previous one
      selectedROIs[0] = combinedROIIndices[roiLabel - 1];
      selectedROIs[1] = i;
      roiManager("Select", selectedROIs);
      roiManager("Combine");
      roiManager("Add", "ffff00");
      combinedROIIndices[roiLabel - 1] = roiManager("index"); //roiManager("index") doesn't work this way
    }
  }
  //ROI manager contains both original fragments and combined fragments---keep only combined ones
  nROIManagerEntries = roiManager("count");
  i = offset;
  do {
    if (!isMember(i, combinedROIIndices, nVoronoiLabels)) {
      roiManager("select", i);
      roiManager("Delete");
      nROIManagerEntries--;
    }
    else i++;
  } while ( i < nROIManagerEntries );
}

function isMember(target, array, arraySize) {
  flag = false;
  i = 0;
  do {
    if (array[i] == target) flag = true;
    i++;
  } while ( flag == false && i < arraySize );
  return flag;
}

function renameROIs(ROIType, offset, nVoronoi, arrayIndices) {
  //offset allows skipping first part of ROI manager list
  for (i = 0; i < nVoronoi; i++) {
    arrayIndices[i] = offset + i;
    VoronoiLabel = getResult("Mode", i);
    roiManager("Select", offset + i); //Circle or ring label to be renamed
    roiManager("Rename", leftPad(ROIType, VoronoiLabel, nVoronoi));
  }
  roiManager("Sort"); //Sort ROI manager list by label
}

function cleanup(message, ch1Title, ch2Title) {
  print(ch1Title + " " + message);
  if (!stopForAnalysisReview) {
    if (isOpen(ch1Title)) {
      selectWindow(ch1Title);
      close();
    }
    if (isOpen(ch2Title)) {
      selectWindow(ch2Title);
      close();
    }
  }
  if (isOpen("Circle")) {
    selectWindow("Circle");
    close();
  }
  if (isOpen("Ring")) {
    selectWindow("Ring");
    close();
  }
  if (isOpen("Voronoi")) {
    selectWindow("Voronoi");
    close();
  }
}
