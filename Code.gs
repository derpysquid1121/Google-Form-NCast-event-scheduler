//Load Moment Library
var moment = Moment.load();
var calChoice = "/*Calendar Here*/";

var GLOBAL = {
  formId : "/*google form here*/",
  calendarVar :"This will be set in the switch",
  calendarId : "/*calendar id here*/",
  calendarId2 : "/*calendar id here*/",
  formMap : {
    eventTitle : "Event Title",
    startDate : "Event Date",
    eventStartTime: "Event Start Time",
    eventEndtime: "Event End Time",
    roomNumber: "Room Number",
    description: "description",
    ipAddress: "IP Address",
  },
}

function onFormSubmit(){
  var eventObject = getFormResponse();
  
  var event = createCalendarEvent(eventObject);
}

function getFormResponse(){
  var form = FormApp.openById(GLOBAL.formId),
      //Create an array of responses
      responses = form.getResponses(),
      //find length of the responses array
      length = responses.length,
      //find the index of the most recent form response
      //since arrays are zero indexed, the last response 
      //is the total number of responses minus one
      lastResponse = responses[length-1] ,
      //get an array of responses to every question item 
      //within the form
      itemResponses = lastResponse.getItemResponses(),
      //create an empty object to store data from the
      //last form response
      eventObject = {};
  for (var i = 0, x = itemResponses.length; i<x; i++) {
    //Get the title of the form item being iterated on
    var thisItem = itemResponses[i].getItem().getTitle(),
        //get the submitted response to the form item being
        //iterated on
        thisResponse = itemResponses[i].getResponse();
    //based on the form question title, map the response of the 
    //item being iterated on into our eventObject variable
    //use the GLOBAL variable formMap sub object to match 
    //form question titles to property keys in the event object
    
    if(i == 0){
      eventObject.eventTitle = thisResponse;
    }
    else if(i==1){
      eventObject.eventStartTime = thisResponse;
    }
    else if(i==2){
      eventObject.eventEndTime = eventObject.eventStartTime.substring(0,11);
      eventObject.eventEndTime = eventObject.eventEndTime.concat(thisResponse);
    }
    else if(i==3){
      eventObject.roomNumber = thisResponse;
      switch (eventObject.roomNumber) {
        case "2211":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "2260":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "Event";
          eventObject.eventTitle =  eventObject.eventTitle;
          calChoice = "/*alternate calendar here*/";
          break;
        case "3250":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/"; 
          //eventObject.description = "G3;V1;C3;R1";
          //eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          eventObject.eventTitle = eventObject.eventTitle;
          calChoice = "/*alternate calendar here*/" 
          break;
        case "5229":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          //eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = eventObject.eventTitle;
          calChoice = "/*alternate calendar here*/" 
          break;
        case "5240":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "5246":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "Lubar":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "G7;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "2225":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "3226":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "3247":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "3253":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "3260":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "3261":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "3268":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "5223":
          eventObject.ipAddress = "/*ip address or room name for room camera here*/";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
      }      
    }
  }
  return eventObject;
}      

function createCalendarEvent(eventObject){
  
  //This is code copied from stack overflow-----------------------Used to try to deal with double form submission------------------------------------
      SpreadsheetApp.flush();                                       // Currently not fixing the problem 
     var lock = LockService.getScriptLock();
  try {
    lock.waitLock(1100); // wait 15 seconds (15000) for others' use of the code section and lock to stop and then proceed
     } catch (e) {
        Logger.log('Could not obtain lock after 30 seconds.');
        return HtmlService.createHtmlOutput("<b> Server Busy please try after some time <p>")
        // In case this a server side code called asynchronously you return a error code and display the appropriate message on the client side
        return "Error: Server busy try again later... Sorry :("
     }
  //--------------------------------------------------------------------------------------------------------------------------------------------------
  //Get a calendar object by opening the calendar using the
  //calendar id stored in the GLOBAL variable object
  //finalCalId = GLOBAL.calendarId;
  var calendar = CalendarApp.getCalendarById(calChoice),
      //The title for the event that will be created
      title = eventObject.eventTitle,
      //The start time and date of the event that will be created
      //For Daylight savings: .subtract(1, 'hours') <-----Paste before .toDate() in two instances below
      startTime = moment(eventObject.eventStartTime).toDate(),//put date and time together
      //The end time and date of the event that will be created
      endTime = moment(eventObject.eventEndTime).toDate(); //put date and time together
      
  //an options object containing the description and guest list
  //for the event that will be created
  var options = {
    description : eventObject.description,//This depends on room number 
    location : eventObject.ipAddress, //This depends on room number
  };
  var event = calendar.createEvent(title, startTime, endTime, options);
  
  return event;
}
  
