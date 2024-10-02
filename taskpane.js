Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        // Initialize the add-in
    }
});

function setOutOfOffice() {
    const scheduleType = document.getElementById('scheduleType').value;
    const startDate = new Date(document.getElementById('startDate').value);
    const endDate = new Date(document.getElementById('endDate').value);
    const message = document.getElementById('message').value;

    if (startDate && endDate && message) {
        switch (scheduleType) {
            case 'oneTime':
                setOneTimeOutOfOffice(startDate, endDate, message);
                break;
            case 'weekly':
                setWeeklyOutOfOffice(startDate, endDate, message);
                break;
            case 'biweekly':
                setBiweeklyOutOfOffice(startDate, endDate, message);
                break;
            default:
                console.error('Invalid schedule type');
        }
    } else {
        console.error('Please fill out all fields.');
        // You can add UI feedback for incomplete form here
    }
}

function setOneTimeOutOfOffice(startDate, endDate, message) {
    Office.context.mailbox.userProfile.getTimeZoneAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const timeZone = asyncResult.value;
            
            Office.context.mailbox.userProfile.setOutOfOfficeSettingsAsync({
                outOfOfficeState: "Scheduled",
                startTime: startDate,
                endTime: endDate,
                internalReply: message,
                externalReply: message,
                timeZone: timeZone
            }, handleAsyncResult);
        } else {
            console.error('Error getting time zone:', asyncResult.error.message);
        }
    });
}

function setWeeklyOutOfOffice(startDate, endDate, message) {
    const recurrence = {
        recurrenceType: "Weekly",
        seriesTime: {
            start: { timeZone: "UTC", dateTime: startDate.toISOString() },
            end: { timeZone: "UTC", dateTime: endDate.toISOString() }
        },
        recurrenceTimeZone: { name: "UTC" },
        recurrenceProperties: { interval: 1, daysOfWeek: ["Monday"] }
    };
    
    setRecurringOutOfOffice(recurrence, message);
}

function setBiweeklyOutOfOffice(startDate, endDate, message) {
    const recurrence = {
        recurrenceType: "Weekly",
        seriesTime: {
            start: { timeZone: "UTC", dateTime: startDate.toISOString() },
            end: { timeZone: "UTC", dateTime: endDate.toISOString() }
        },
        recurrenceTimeZone: { name: "UTC" },
        recurrenceProperties: { interval: 2, daysOfWeek: ["Monday"] }
    };
    
    setRecurringOutOfOffice(recurrence, message);
}

function setRecurringOutOfOffice(recurrence, message) {
    Office.context.mailbox.makeEwsRequestAsync(
        `<?xml version="1.0" encoding="utf-8"?>
        <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                       xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
                       xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
                       xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
          <soap:Header>
            <t:RequestServerVersion Version="Exchange2013" />
          </soap:Header>
          <soap:Body>
            <m:CreateItem>
              <m:Items>
                <t:CalendarItem>
                  <t:Subject>Out of Office</t:Subject>
                  <t:Body BodyType="HTML">${message}</t:Body>
                  <t:Start>${recurrence.seriesTime.start.dateTime}</t:Start>
                  <t:End>${recurrence.seriesTime.end.dateTime}</t:End>
                  <t:Recurrence>
                    <t:${recurrence.recurrenceType}Recurrence>
                      <t:Interval>${recurrence.recurrenceProperties.interval}</t:Interval>
                      <t:DaysOfWeek>${recurrence.recurrenceProperties.daysOfWeek[0]}</t:DaysOfWeek>
                    </t:${recurrence.recurrenceType}Recurrence>
                  </t:Recurrence>
                </t:CalendarItem>
              </m:Items>
            </m:CreateItem>
          </soap:Body>
        </soap:Envelope>`,
        handleAsyncResult
    );
}

function handleAsyncResult(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log('Out of office settings updated successfully.');
        // You can add UI feedback here
    } else {
        console.error('Error updating out of office settings:', asyncResult.error.message);
        // You can add error handling UI here
    }
}

