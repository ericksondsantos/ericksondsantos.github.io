const flames = ['Friends', 'Lovers', 'Affection', 'Marriage', 'Enemies', 'Soulmates'];
const yourName = document.getElementById('yourName');
const theirName = document.getElementById('theirName');
const feedback = document.getElementById('feedback');

function checkAnswer(event) {
    event.preventDefault(); // Prevents page reload on form submit
    const name1 = new Set(yourName.value.trim().toLowerCase().replace(/[^a-z]/g, ""));
    const name2 = new Set(theirName.value.trim().toLowerCase().replace(/[^a-z]/g, ""));
    let mismatches = 0;
    for (let char of name1) {
        if (!name2.has(char)) {
        mismatches++;
        }
    }
    for (let char of name2) {
        if (!name1.has(char)) {
        mismatches++;
        }
    }
    if (mismatches === 0) {
        feedback.textContent = 'Perfect Match!';
        return;
    }
    const flameIndex = mismatches > 0 ? (mismatches - 1) % flames.length : 0;
    feedback.textContent = `${flames[flameIndex]}!`;
}
