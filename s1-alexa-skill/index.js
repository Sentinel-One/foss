'use strict';

var waitUntil = require('wait-until');
var request = require('request');
var timestamp = require('unix-timestamp');

var server = '';
var host = '';
var apiKey = '';

// Main event router 
exports.handler = function (event, context) {
    try {
        console.log("event.session.application.applicationId=" + event.session.application.applicationId);

        if (event.session.application.applicationId !== "amzn1.ask.skill.d3705301-36ac-40c4-b3b0-399279bcd883") {
             context.fail("Invalid Application ID");
        }

        if (event.session.new) {
            onSessionStarted({requestId: event.request.requestId}, event.session); }

        if (event.session.user.accessToken === undefined) {
            console.log('No access token, need to link');
            onLink(event.request, event.session,
                function callback(sessionAttributes, speechletResponse) {
                    context.succeed(buildResponse(sessionAttributes, speechletResponse)); }); 
        } else if (event.request.type === "LaunchRequest") {
            onLaunch(event.request, event.session,
                function callback(sessionAttributes, speechletResponse) {
                    context.succeed(buildResponse(sessionAttributes, speechletResponse)); });
        } else if (event.request.type === "IntentRequest") {
            var token = event.session.user.accessToken;
            host = 'https://' + token.substring(0, token.indexOf('|'));
            server = token.substring(0, token.indexOf('|'));
            apiKey = 'Token ' +token.substring(token.indexOf('|')+1);
            onIntent(event.request, event.session,
                function callback(sessionAttributes, speechletResponse) {
                    context.succeed(buildResponse(sessionAttributes, speechletResponse)); });
        } else if (event.request.type === "SessionEndedRequest") {
            onSessionEnded(event.request, event.session);
            context.succeed();
        }
    } catch (e) { context.fail("Exception: " + e); }
};

/**
 * Called when the session starts.
 */
function onSessionStarted(sessionStartedRequest, session) {
    console.log("onSessionStarted requestId=" + sessionStartedRequest.requestId + ", sessionId=" + session.sessionId); }

/**
 * Called when the user invokes the skill without specifying what they want.
 */
function onLaunch(launchRequest, session, callback) {
    getWelcomeResponse(callback);
    /*
    console.log("onLaunch requestId=" + launchRequest.requestId + ", sessionId=" + session.sessionId);
    var cardTitle = "SentinelOne";
    var speechOutput = "You can tell SentinelOne to say Summary Report or Daily Report";
    var repromptText = "Try saying Summary Report or Daily Report";
    callback(session.attributes, buildSpeechletResponse(cardTitle, speechOutput, repromptText, false));
    */
}

/**
 * Called when linking to account
 */
function onLink(linkRequest, session, callback) {
    getLinkResponse(callback);
}

/**
 * Called when the user specifies an intent for this skill.
 */
function onIntent(intentRequest, session, callback) {
    console.log("onIntent requestId=" + intentRequest.requestId + ", sessionId=" + session.sessionId);

    var intent = intentRequest.intent,
        intentName = intentRequest.intent.name;

    // dispatch custom intents to handlers here
    if (intentName === 'SummaryReport')
        handleSummaryRequest(intent, session, callback);
    else if (intentName === 'DailyReport')
        handleDailyRequest(intent, session, callback);
    else if (intentName === 'ServerInfo')
        handleServerInfoRequest(intent, session, callback);
    else if (intentName === "AMAZON.HelpIntent")
       getHelpResponse(callback); 
    else if (intentName === "AMAZON.StopIntent" || intentName === "AMAZON.CancelIntent")
       handleSessionEndRequest(callback);     
    else {
        throw "Invalid intent";
    }
}

/**
 * Called when the user ends the session.
 * Is not called when the skill returns shouldEndSession=true.
 */
function onSessionEnded(sessionEndedRequest, session) {
    console.log("onSessionEnded requestId=" + sessionEndedRequest.requestId + ", sessionId=" + session.sessionId);
}

function getHelpResponse(callback) { 
    var sessionAttributes = {}; 
    var cardTitle = "SentinelOne"; 
    var speechOutput = "Try saying Summary Report or Daily Report";
    var repromptText = "Just say Summary Report or Daily Report";
    callback(sessionAttributes, buildSpeechletResponse(cardTitle, speechOutput, repromptText, false)); 
} 

function getWelcomeResponse(callback) { 
    var sessionAttributes = {}; 
    var cardTitle = "SentinelOne"; 
    var speechOutput = "Welcome! You can ask SentinelOne to say Summary Report or Daily Report";
    var repromptText = "Try saying Summary Report or Daily Report";
    callback(sessionAttributes, buildSpeechletResponse(cardTitle, speechOutput, repromptText, false)); 
} 

function getLinkResponse(callback) { 
    var sessionAttributes = {}; 
    var cardTitle = "SentinelOne"; 
    var speechOutput = "Please link this account to your SentinelOne server";
    callback(sessionAttributes, buildSpeechletResponseToLink(cardTitle, speechOutput, null, true)); 
} 

function handleSessionEndRequest(callback) { 
    var sessionAttributes = {}; 
    var cardTitle = "Session Ended"; 
    var speechOutput = "Thank you for using the SentinelOne skill"; 
    callback(sessionAttributes, buildSpeechletResponse(cardTitle, speechOutput, null, true)); 
} 

// Summary Report --------------------------------------------------------------

function urlThreatSummary() { 
        return {
            url: host + '/web/api/v1.6/threats/summary',
            headers: {
                'Authorization' : apiKey,
                'Content-Type' : 'application/json'
            } 
        };
}

function urlEndpointSummary() { 
        return {
            url: host + '/web/api/v1.6/agents/count-by-filters?participating_fields=hardware_information__machine_type,is_up_to_date,software_information__os_type,infected',
            headers: {
                'Authorization' : apiKey,
                'Content-Type' : 'application/json'
            } 
        };
}

function getSummaryJSON(callback) {

    var textOut = '<speak>';
    var textThreat = '';
    var textEndpoint = '';
    var waiting = 2;
    
    request.get(urlThreatSummary(), function(error, response, body) {
                console.log('After Summary');

        var d = JSON.parse(body);
        
        textThreat += "<p>Let's start with the threat status</p>";
        var active = d['active'];
        if (active === undefined) {
            textThreat += '<p>Sorry there is an error</p>';
        } else if (active === 0) {
            textThreat += '<p>Congratulations, there are currently no active and unresolved threats</p>';
        } else if (active === 1)
        {
            textThreat += '<p>There is currently ' + active + ' active threat that has not been resolved</p>' +
                          '<p>This is important and should be investigated immediately</p>';
        } else {
            textThreat += '<p>There are currently ' + active + ' active threats that have not been resolved</p>' +
                          '<p>This is important and should be investigated immediately</p>';
        }

        var mitigated = d['mitigated'];
        var blocked = d['blocked'];
        var takenCare = mitigated + blocked;
        if (takenCare === undefined) {
            textThreat += '<p>Sorry there is an error</p>';
        } else if (takenCare === 0) {
            // There are no unresolved Mitigated threats, so nothing to say here
        } else if (takenCare === 1)
        {
            textThreat += '<p>One threat was automatically prevented</p><p>it should be verified and acknowledged in the console</p>';
        } else {
            textThreat += '<p>' + takenCare + ' threats are automatically prevented</p><p>they should be verified and acknowledged in the console</p>';
        }

        var suspicious = d['suspicious'];
        if (suspicious === undefined) {
            textThreat += '<p>Sorry there is an error</p>';
        } else if (suspicious === 0) {
            // There are no unresolved Mitigated threats, so nothing to say here
        } else if (suspicious === 1)
        {
            textThreat += '<p>In addition</p><p>One suspicious activity was detected</p><p>it should be investigated</p>';
        } else {
            textThreat += '<p>In addition</p><p>' + suspicious + ' suspicious activities were detected</p><p>they should be investigated</p>';
        }
        
        textThreat += '<break time="1500ms"/><p>Next up the summary of installed endpoints</p>';
        --waiting;
    });

    request.get(urlEndpointSummary(), function(error, response, body) {
        var d = JSON.parse(body);
        
        var count = d['total_count'];
        if (count === undefined) {
            textEndpoint += '<p>Sorry there is an error</p>';
        } else if (count === 0) {
            textEndpoint += '<p>There are currently no endpoints installed for this SentinelOne server</p>';
        } else if (count === 1)
        {
            textEndpoint += '<p>There is currently ' + count + ' endpoint installed</p>';
        } else {
            textEndpoint += '<p>There are currently ' + count + ' endpoints installed</p>';
        }
        
        var win = (d['software_information__os_type']['windows'] === undefined ? 0 : d['software_information__os_type']['windows']) + (d['software_information__os_type']['windows_legacy'] === undefined ? 0 : d['software_information__os_type']['windows_legacy']);
        var mac = (d['software_information__os_type']['osx'] === undefined ? 0 : d['software_information__os_type']['osx']);
        var linux = (d['software_information__os_type']['linux'] === undefined ? 0 : d['software_information__os_type']['linux']);
        
        if (win > 0) {
            textEndpoint += '<p>' + Math.round((win/count)*100) + ' percent of the endpoints are Windows</p>';
        } 
        
        if (mac > 0) {
            textEndpoint += '<p>' + Math.round((mac/count)*100) + ' percent of the endpoints are macOS</p>';
        } 

        if (linux > 0) {
            textEndpoint += '<p>' + Math.round((linux/count)*100) + ' percent of the endpoints are Linux</p>';
        } 

        var uptodate = (d['is_up_to_date']['true'] === undefined ? 0 : d['is_up_to_date']['true']);
        if (uptodate > 0) {
            textEndpoint += '<p>' + Math.round((uptodate/count)*100) + ' percent of the endpoints are up to date with the latest version of the agent</p>';
        } 

        var infected = (d['infected']['true'] === undefined ? 0 : d['infected']['true']);
        if (infected === 0)
            textEndpoint += '<p>Lastly, good news that <emphasis level="strong">none</emphasis> of the endpoints are currently infected with threats</p>';
        else if (infected === 1) {
            textEndpoint += '<p>Lastly, ' + infected + ' endpoint is currently infected with threats, which is ' + Math.round((infected/count)*100) + ' percent of the total population</p>';
        } else if (infected > 0) {
            textEndpoint += '<p>Lastly, ' + infected + ' endpoints are currently infected with threats, which are ' + Math.round((infected/count)*100) + ' percent of the total population</p>';
        } 

        --waiting;
    });
    
    waitUntil()
        .interval(100)
        .times(30)
        .condition(function() {
            return (waiting === 0 ? true : false)
        })
        .done (function (result) {
            textOut += textThreat + textEndpoint + "<break time='1500ms'/><p>That's all, please have a great day!</p>" + '</speak>';
            callback(textOut);
        });
} 

function handleSummaryRequest(intent, session, callback) {
    var speechOutput = "We have an error";
    getSummaryJSON(function (data) {
        if (data != "ERROR") {
            speechOutput = data;
        }
        callback(session.attributes, buildSpeechletResponseWithoutCard(speechOutput, "", "true"));
        
    });
}

// Daily Report ----------------------------------------------------------------

function urlDailySummary() { 
    
        var today = timestamp.now().toString().replace('.','');
        var yesterday = timestamp.now('-1d').toString().replace('.','');
        return {
            url: host + '/web/api/v1.6/threats?limit=1000&skip=0&created_at__lte=' + today +  '&created_at__gte=' + yesterday,
            headers: {
                'Authorization' : apiKey,
                'Content-Type' : 'application/json'
            } 
        };
}

function getDailyJSON(callback) {

    var textOut = '<speak>';
    var textThreat = '';
    var textEndpoint = '';
    var waiting = 1;
    
    request.get(urlDailySummary(), function(error, response, body) {
        var d = JSON.parse(body);
        
        textThreat += "<p>Here is the SentinelOne daily report</p>";
        
        var numberThreat = d.length;

        if (numberThreat === 0) {
            textThreat += '<p>Good news, there are no threats detected in the last 24 hours</p>';
        } else {
            if (numberThreat === 1)
            {
                textThreat += '<p>There is ' + numberThreat + ' threat that has been detected in the last 24 hours</p>';
            }
            else
            {
                textThreat += '<p>There are ' + numberThreat + ' threats that have been detected in the last 24 hours</p>';
            }

            var CountActive = 0;
            var CountPrevented = 0;
            var CountSuspicious = 0;
            for (var i = 0; i < d.length; i++) {
                var threat = d[i];
                if (threat.resolved === false && threat.mitigation_status === 1)
                {
                    CountActive++;
                }
                else if (threat.resolved === false && (threat.mitigation_status === 0 || threat.mitigation_status === 2))
                {
                    CountPrevented++;
                }
                else if (threat.resolved === false && threat.mitigation_status === 3)
                {
                    CountSuspicious++;
                }
            }
            
            /*
                0 - Mitigated
                1 - Active
                2 - Blocked
                3 - Suspicious
                4 - Pending
                5 - Suspicious canceled
            */
            
            if (CountActive === 0) {
                textThreat += '<p>Good news, none of the threats are still active and all have been resolved.</p>';
            } else if (CountActive === 1)
            {
                textThreat += '<p>There is currently ' + CountActive + ' active threat that has not been resolved</p>' +
                              '<p>This is important and should be investigated immediately</p>';
            } else {
                textThreat += '<p>There are currently ' + CountActive + ' active threats that have not been resolved</p>' +
                              '<p>These are the urgent items that should be investigated immediately</p>';
            }
            
            if (CountPrevented === 0) {
                // There are no unresolved Mitigated threats, so nothing to say here
            } else if (CountPrevented === 1)
            {
                textThreat += '<p>One threat was automatically prevented</p><p>just acknowledge it in the console, and you are all set</p>';
            } else {
                textThreat += '<p>' + CountPrevented + ' threats are automatically prevented</p><p>just acknowledge them in the console, and you are all set</p>';
            }            

            if (CountSuspicious === 0) {
                // There are no unresolved Mitigated threats, so nothing to say here
            } else if (CountSuspicious === 1)
            {
                textThreat += '<p>Lastly</p><p>one suspicious activity was detected</p><p>it should be investigated</p>';
            } else {
                textThreat += '<p>Lastly</p><p>' + CountSuspicious + ' suspicious activities were detected</p><p>they should be investigated</p>';
            }

        }
        --waiting;
    });

    waitUntil()
        .interval(500)
        .times(30)
        .condition(function() {
            return (waiting === 0 ? true : false);
        })
        .done (function (result) {
            textOut += textThreat + textEndpoint + "<break time='1500ms'/><p>That's all for the daily report.</p>" + '</speak>';
            callback(textOut);
        });
} 

function handleDailyRequest(intent, session, callback) {
    var speechOutput = "We have an error";
    getDailyJSON(function (data) {
        if (data != "ERROR") {
            speechOutput = data;
        }
        callback(session.attributes, buildSpeechletResponseWithoutCard(speechOutput, "", "true"));
    });
}

// Server Information ----------------------------------------------------------------

function urlServerInfoSummary() { 
        return {
            url: host + '/web/api/v1.6/info',
            headers: {
                'Authorization' : apiKey,
                'Content-Type' : 'application/json'
            } 
        };
}

function getServerInfoJSON(callback) {

    var textOut = '<speak>';
    var textInfo = '';
    var waiting = 1;
    
    request.get(urlServerInfoSummary(), function(error, response, body) {
        var d = JSON.parse(body);
        
        textInfo += "<p>Here is the information about your SentinelOne server</p>";
        
        var version = d.version + '.' + d.build;

        textInfo += '<p>The name of your server is ' + server + ' which is spelt <prosody rate="x-slow"><say-as interpret-as="spell-out">' + server.substring(0, server.indexOf('.')) + '</say-as></prosody></p>';
        textInfo += '<p>It is of version <emphasis level="strong">' + version + '</emphasis></p>';
        --waiting;
    });

    waitUntil()
        .interval(100)
        .times(30)
        .condition(function() {
            return (waiting === 0 ? true : false);
        })
        .done (function (result) {
            textOut += textInfo + "<break time='1500ms'/><p>That's all on server information!</p>" + '</speak>';
            callback(textOut);
        });
} 

function handleServerInfoRequest(intent, session, callback) {
    var speechOutput = "We have an error";
    getServerInfoJSON(function (data) {
        if (data != "ERROR") {
            speechOutput = data;
        }
        callback(session.attributes, buildSpeechletResponseWithoutCard(speechOutput, "", "true"));
    });
}

// ------- Helper functions to build responses -------

function buildSpeechletResponse(title, output, repromptText, shouldEndSession) {
    
    return {
        outputSpeech: {
            type: "PlainText",
            text: output
        },
        card: {
            type: "Simple",
            title: title,
            content: output
        },
        reprompt: {
            outputSpeech: {
                type: "PlainText",
                text: repromptText
            }
        },
        shouldEndSession: shouldEndSession
    };
}

function buildSpeechletResponseToLink(title, output, repromptText, shouldEndSession) {
    
    return {
        outputSpeech: {
            type: "PlainText",
            text: "Please use the companion app to link to your SentinelOne server"
        },
        card: {
            type: "LinkAccount",
            title: title,
            content: output
        },
        reprompt: {
            outputSpeech: {
                type: "PlainText",
                text: repromptText
            }
        },
        shouldEndSession: shouldEndSession
    };
}

function buildSpeechletResponseWithoutCard(output, repromptText, shouldEndSession) {
    return {
        outputSpeech: {
            type: "SSML",
            ssml: output
        },
        reprompt: {
            outputSpeech: {
                type: "PlainText",
                text: repromptText
            }
        },
        shouldEndSession: shouldEndSession
    };
}

function buildResponse(sessionAttributes, speechletResponse) {
    return {
        version: "1.0",
        sessionAttributes: sessionAttributes,
        response: speechletResponse
    };
}
