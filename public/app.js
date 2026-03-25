async function loadBootstrap() {
  const output = document.getElementById("bootstrap-output");

  try {
    const response = await fetch("/api/bootstrap");
    if (!response.ok) {
      throw new Error(`Bootstrap request failed with ${response.status}`);
    }

    const data = await response.json();
    output.textContent = JSON.stringify(data, null, 2);
  } catch (error) {
    output.textContent = `Bootstrap failed: ${error.message}`;
  }
}

loadBootstrap();
