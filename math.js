        // Game State Variables
        let currentAnswer = 0;
        let score = 0;
        let streak = 0;

        // DOM Elements
        const equationEl = document.getElementById('equation');
        const userAnswerEl = document.getElementById('user-answer');
        const feedbackEl = document.getElementById('feedback');
        const scoreEl = document.getElementById('score');
        const streakEl = document.getElementById('streak');

        // Helper function to get random integer
        function getRandomInt(min, max) {
            return Math.floor(Math.random() * (max - min + 1)) + min;
        }

        // Generate a new random math equation
        function generateEquation() {
            // Clear previous feedback and input
            feedbackEl.textContent = '';
            userAnswerEl.value = '';
            userAnswerEl.focus();

            const operators = ['+', '-', '×', '÷'];
            const operator = operators[Math.floor(Math.random() * operators.length)];
            
            let num1, num2;

            switch(operator) {
                case '+':
                    num1 = getRandomInt(5, 100);
                    num2 = getRandomInt(5, 100);
                    currentAnswer = num1 + num2;
                    break;
                case '-':
                    // Ensure results aren't frustratingly negative for a casual game
                    num1 = getRandomInt(10, 100);
                    num2 = getRandomInt(5, num1); 
                    currentAnswer = num1 - num2;
                    break;
                case '×':
                    num1 = getRandomInt(2, 12);
                    num2 = getRandomInt(2, 12);
                    currentAnswer = num1 * num2;
                    break;
                case '÷':
                    // Prevent decimals: Generate answer and num2 first, then multiply to get num1
                    num2 = getRandomInt(2, 10);
                    currentAnswer = getRandomInt(2, 10);
                    num1 = num2 * currentAnswer;
                    break;
            }

            // Display the equation to the user
            equationEl.textContent = `${num1} ${operator} ${num2}`;
        }

        // Check if the user's answer is correct
        function checkAnswer(event) {
            event.preventDefault(); // Prevents page reload on form submit

            // If an equation hasn't been generated yet
            if (equationEl.textContent === '---') {
                feedbackEl.textContent = "Click 'Randomize' first!";
                feedbackEl.className = "feedback incorrect";
                return;
            }

            const userAnswer = parseInt(userAnswerEl.value, 10);

            if (userAnswer === currentAnswer) {
                feedbackEl.textContent = "Good job!";
                feedbackEl.className = "feedback correct";
                score += 1;
                streak++;
                scoreEl.textContent = score;
                streakEl.textContent = streak;
                // Automatically auto-advance to next question after a brief delay
                setTimeout(generateEquation, 2400);
                createStar();
            } else {
                feedbackEl.textContent = `Try again!`;
                feedbackEl.className = "feedback incorrect";
                score -= 1;
                streak = 0;
                scoreEl.textContent = score;
                streakEl.textContent = streak;
                userAnswerEl.select(); // Highlight user input for easy correction
            }
        }

        function createStar() {
            const star = document.createElement('div');
            star.classList.add('star');

            // Determine a random size between 2px and 5px
            const size = Math.random() * 3 + 2;
            star.style.width = `${size}px`;
            star.style.height = `${size}px`;

            // Pick a random X and Y coordinate based on the window's dimensions
            const randomX = Math.random() * window.innerWidth;
            const randomY = Math.random() * window.innerHeight;
            star.style.left = `${randomX}px`;
            star.style.top = `${randomY}px`;

            // Give it a subtle random color variation (white, light blue, or soft yellow)
            const colors = ['#ffffff', '#ffffff', '#e0f2fe', '#fef08a'];
            const randomColor = colors[Math.floor(Math.random() * colors.length)];
            star.style.backgroundColor = randomColor;
            
            // Add a soft glow that matches its color
            star.style.boxShadow = `0 0 ${size * 2}px ${randomColor}`;

            // Append the newly created star to the body
            document.body.appendChild(star);
        }

        // Initialize the first question when the page loads
        window.onload = generateEquation;