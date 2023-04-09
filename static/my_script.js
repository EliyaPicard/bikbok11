// Get all "Show More" and "Show Less" buttons
const showMoreButtons = document.querySelectorAll(".show-more");
const showLessButtons = document.querySelectorAll(".show-less");

// Loop through all "Show More" buttons and add click event listener
showMoreButtons.forEach((button) => {
  button.addEventListener("click", () => {
    // Get the hidden span element containing the rest of the wine type
    const hiddenSpan = button.parentElement.querySelector(".hidden");
    // Show the hidden span
    hiddenSpan.style.display = "inline";
    // Hide the "Show More" button
    button.style.display = "none";
  });
});

// Loop through all "Show Less" buttons and add click event listener
showLessButtons.forEach((button) => {
  button.addEventListener("click", () => {
    // Get the hidden span element containing the rest of the wine type
    const hiddenSpan = button.parentElement.querySelector(".hidden");
    // Hide the hidden span
    hiddenSpan.style.display = "none";
    // Show the "Show More" button
    button.parentElement.querySelector(".show-more").style.display = "inline";
  });
});
