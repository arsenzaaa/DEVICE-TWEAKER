// Guide filter: the search field and [ TAG ] chips work like the device filter of the program.
(function () {
  var bar = document.querySelector(".filterbar");
  if (!bar) return;
  var input = bar.querySelector(".search");
  var chips = bar.querySelectorAll(".chip");
  var cards = document.querySelectorAll(".cards .card");
  var empty = document.querySelector(".empty");
  var tag = "*";

  function apply() {
    var query = input.value.trim().toLowerCase();
    var shown = 0;
    cards.forEach(function (card) {
      var match = (tag === "*" || card.dataset.tag === tag) && (!query || card.dataset.text.indexOf(query) !== -1);
      card.hidden = !match;
      if (match) shown++;
    });
    if (empty) empty.hidden = shown !== 0;
  }

  chips.forEach(function (chip) {
    chip.addEventListener("click", function () {
      tag = chip.dataset.filter;
      chips.forEach(function (c) { c.classList.toggle("on", c === chip); });
      apply();
    });
  });
  input.addEventListener("input", apply);
  document.addEventListener("keydown", function (e) {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "f") {
      e.preventDefault();
      input.focus();
      input.select();
    }
  });
})();
