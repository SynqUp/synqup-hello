/* global Office */
Office.onReady(() => {});

Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);

async function onAppointmentSendHandler(event) {
  try {
    event.completed({
      allowEvent: false,
      errorMessage: "Hello from SynqUp - add-in is installed"
    });
  } catch (e) {
    event.completed({ allowEvent: true });
  }
}
