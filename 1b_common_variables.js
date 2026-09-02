/**
These are variables that are used among multiple different files, mainly 2a and 3a
*/


// the main sheet tab with all the buttons for running scripts
const BASE_SHEET_NAME = '[Base]';


// helper constants used for setting FROM and TO times
const NOW = new Date().toISOString().split('T')[0]; // Today (YYYY-MM-DD)
const ONE_DAY_AGO = daysAgo(1);
const THREE_DAYS_AGO = daysAgo(3);
const ONE_WEEK_AGO = daysAgo(7);
const ONE_MONTH_AGO = daysAgo(30);
const TWO_MONTHS_AGO = daysAgo(60);
const THREE_MONTHS_AGO = daysAgo(90);
const FOUR_MONTHS_AGO = daysAgo(120);
const FIVE_MONTHS_AGO = daysAgo(150);
const SIX_MONTHS_AGO = daysAgo(180);


// request constants, these are sent to Zoom API as part of the request
// time period of past meetings to GET
const FROM = ONE_MONTH_AGO;
const TO = NOW;
// Type of meeting (meeting or webinar, can also send "" for both)
const MEETING_TYPE = "meeting";
// Optional search query key if you only want meetings with specific word(s) in the topic name
const SEARCH_KEY = "";
// Max meetings per request page (up to 300), used to lower request rate to prevent hitting API rate limits
const PAGE_SIZE = 200;


// toggle to include only meetings that occurr on 4th thursdays of the month
const ONLY_FOURTH_THURS = true;
