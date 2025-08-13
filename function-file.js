/* global Office */
Office.onReady(() => {
  console.log("SynqUp build 2025-08-13-1"); // <-- change this each time you publish
});

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
