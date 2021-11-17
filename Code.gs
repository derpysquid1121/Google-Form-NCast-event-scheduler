//Load Moment Library
var moment = Moment.load();
var calChoice = "uwlawschoolhelp@gmail.com";

var GLOBAL = {
  formId : "1Z6sopaaMSld4ZLP9Gt1mlZvSjfMl-2RUPePn5cAm9OM",
  calendarVar :"This will be set in the switch",
  calendarId : "uwlawschoolhelp@gmail.com"/*"2hs6j46a9vqs1n1fa3h63iblk0@group.calendar.google.com"*/,
  calendarId2 : "b7gm6m86e11r29dg2osa1gbalk@group.calendar.google.com",
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
      //eventObject.eventEndTime = eventObject.eventEndTime;
      //debugger;
      /*
      eventObject.eventEndTime = moment(eventObject.eventStartTime);//Take start date and tack on second time.
      eventObject.eventEndTime.add(parseInt(thisResponse), 'h').format('YYYY-MM-DD hh:mm:ss');
      */
    }
    else if(i==3){
      eventObject.roomNumber = thisResponse;
      switch (eventObject.roomNumber) {
        case "2211":
          eventObject.ipAddress = "128.104.94.137";
          eventObject.description = "C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "2260":
          eventObject.ipAddress = "Law2260";
          eventObject.description = "Event";
          eventObject.eventTitle =  eventObject.eventTitle;
          calChoice = "b7gm6m86e11r29dg2osa1gbalk@group.calendar.google.com";
          break;
        case "3250":
          eventObject.ipAddress = "Law3250"; //"128.104.94.146"
          //eventObject.description = "G3;V1;C3;R1";
          //eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          eventObject.eventTitle = eventObject.eventTitle;
          calChoice = "b7gm6m86e11r29dg2osa1gbalk@group.calendar.google.com" 
          break;
        case "5229":
          eventObject.ipAddress = "Law5229"; //"128.104.94.147"
          //eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = eventObject.eventTitle;
          calChoice = "b7gm6m86e11r29dg2osa1gbalk@group.calendar.google.com"
          break;
        case "5240":
          eventObject.ipAddress = "128.104.94.139";
          eventObject.description = "C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "5246":
          eventObject.ipAddress = "128.104.94.140";
          eventObject.description = "C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "Lubar":
          eventObject.ipAddress = "128.104.94.145";
          eventObject.description = "G7;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "2225":
          eventObject.ipAddress = "128.104.94.148";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "3226":
          eventObject.ipAddress = "128.104.94.149";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "3247":
          eventObject.ipAddress = "128.104.94.150";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "3253":
          eventObject.ipAddress = "128.104.94.151";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "3260":
          eventObject.ipAddress = "128.104.94.152";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "3261":
          eventObject.ipAddress = "128.104.94.162";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "3268":
          eventObject.ipAddress = "128.104.94.163";
          eventObject.description = "G3;V1;C3;R1";
          eventObject.eventTitle = "PR720<" + eventObject.eventTitle + ">";
          break;
        case "5223":
          eventObject.ipAddress = "128.104.94.164";
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
  
