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
            case 'onetime':
                setOneTimeOutOfOffice(startDate, endDate, message);
                break;
            case 'weekly':
                setWeeklyOutOfOffice(startDate, endDate, message);
                break;
            case 'biweekly':
                setBiWeeklyOutOfOffice(startDate, endDate, message);
                break;
            default:
                console.error('Invalid schedule type.');
        }
    } else {
        console.error('Please fill out all fields.');

    }
function setOneTimeOutOfOffice(startDate, endDate, message) {
    Office.context.mailbox.useProfile.getTimeZoneAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const timeZone = asyncResult.value;

            Office.context.mailbox.userProfile.setOutOfOfficeSettingsAsync({
                outOfOfficeState: "Scheduled",
                startTime: startDate,
                endTime: endDate,
                internalReply: message,
                externalReply: message,
                timeZone: timeZone,
            }, (handleAsyncResult);
        } else {
            console.error('Error getting time zone:', asyncResult.error.message);
        }
    });
}

function setWeeklyOutOfOffice(startDate, endDate, message) {
    const recurrence = {
        recurrenceType: "Weekly",
        seriesTime: {
            start: { timeZone: "GMT +1", dateTime: startDate.toISOString() },
            end: { timeZone: "GMT +1", dateTime: endDate.toISOString() }
        },
        recurrenceTimeZone: { name: "GMT +1" },
        recurrenceProperties: { interval: 1, daysOfWeekd: ["Monday"] }
    };

    setRecurringOutOfOffice(recurrence, message);
}

function setBiweeklyOutOfOffice(startDate, endDate, message) {
    const recurrence = {
        recurrenceType: "Biweekly",
        seriesTime: {
            start: { timeZone: "GMT +1", dateTime: startDate.toISOString() },
            end: { timeZone: "GMT +1", dateTime: endDate.toISOString() }
        },
        recurrenceTimeZone: { name: "GMT +1" },
        recurrenceProperties: { interval: 2, daysOfWeek: ["Monday"] }
    };

    setRecurringOutOfOffice(recurrence, message);
}
        Office.context.mailbox.userProfile.getTimeZoneAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const timeZone = asyncResult.value;

                Office.context.mailbox.userProfile.setOutOfOfficeSettingsAsync({
                    outOfOfficeState: "Scheduled",
                    startTime: starteDate,
                    endTime: endDate,
                    internalReply: message,
                    externalReply: message,
                    timeZone: timeZone
                }, (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded){
                        console.log("Out of office settings updated successfully.");
                        // You can add UI feedback here if needed
                    } else {
                        console.error('Error updating out of office settings:', asyncResult.error.message);
                        // You can add UI feedback here if needed
                    }
                });
            } else {
                console.error('Error getting time zone:', asyncResult.error.message);
            }
        });
    } else {
        alert('Please fill out all fields.');
        // You can add UI feedback here if needed for imcomplete form here
    }
}
    