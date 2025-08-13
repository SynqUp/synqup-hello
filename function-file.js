/* global Office */
Office.initialize = () => {
  console.log("SynqUp build 2025-08-13-2");
};

function synqupOnSend(event) {
  try {
    event.completed({
      allowEvent: false,
      errorMessage: "Hello from SynqUp - add-in is installed"
    });
  } catch (e) {
    event.completed({ allowEvent: true });
  }
}

if (Office.actions && Office.actions.associate) {
  Office.actions.associate("synqupOnSend", synqupOnSend);
}
