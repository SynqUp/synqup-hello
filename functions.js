/* global Office */
Office.initialize = () => {};

// This runs when the user hits Send on an email or a meeting
function synqupOnSend(event) {
  const item = Office.context.mailbox.item;

  // 1 - agenda check: require the word "Agenda" or at least 20 chars
  item.body.getAsync(Office.CoercionType.Text, asyncResult1 => {
    const bodyText = (asyncResult1.value || "").trim();
    const hasAgenda = /agenda/i.test(bodyText) || bodyText.length >= 20;

    // 2 - participant count check if this is an appointment
    const haveAttendeeApis =
      item.requiredAttendees && item.optionalAttendees &&
      item.requiredAttendees.getAsync && item.optionalAttendees.getAsync;

    if (!haveAttendeeApis) {
      // not an appointment compose - just enforce agenda
      if (!hasAgenda) {
        event.completed({ allowEvent: false, errorMessage: "Please add an agenda." });
      } else {
        event.completed({ allowEvent: true });
      }
      return;
    }

    item.requiredAttendees.getAsync(r1 => {
      item.optionalAttendees.getAsync(r2 => {
        const req = Array.isArray(r1.value) ? r1.value.length : 0;
        const opt = Array.isArray(r2.value) ? r2.value.length : 0;
        const total = req + opt;

        if (!hasAgenda) {
          event.completed({ allowEvent: false, errorMessage: "Please add an agenda." });
        } else if (total > 8) {
          event.completed({ allowEvent: false, errorMessage: "Keep participants under 8 people." });
        } else {
          event.completed({ allowEvent: true });
        }
      });
    });
  });
}

// Map manifest function name to JS function
if (Office.actions && Office.actions.associate) {
  Office.actions.associate("synqupOnSend", synqupOnSend);
}
