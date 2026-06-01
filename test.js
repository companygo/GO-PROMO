async function test() {
  try {
    const res = await fetch("https://script.google.com/macros/s/AKfycbyWNSW__zb0w5rFRmjYvujn5_kt5hUrUmobiXjwXuy84RdRcjJeuwRA6koutt-B6DPxOg/exec?action=winners");
    const text = await res.text();
    console.log("Response:", text);
  } catch (err) {
    console.error("Error:", err);
  }
}
test();
