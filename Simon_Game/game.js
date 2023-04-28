var buttonColors = ["red", "blue", "green", "yellow"];
var gamePattern = [];
var userClickedPattern = [];
var level = 0;
var started = false;

$(document).keydown(function() {
    if(started === false){
        nextSequence();
    }
});

$(document).click(function() {
    if(started === false){
        nextSequence();
    }
});

function nextSequence() {
    started = true;
    $("#level-title").text("Level: " + level);
    var randomNumber = Math.floor(Math.random() * 4);
    var randomChosenColor = buttonColors[randomNumber];
    gamePattern.push(randomChosenColor);
    animateButton(randomChosenColor);
    level++;
}

$(".btn").click(function(event) {
    var userChosenColor = event.target.id;
    userClickedPattern.push(userChosenColor);
    animateButton(userChosenColor);
    if(gamePattern.length === userClickedPattern.length) {
        checkAnswer(level);
    }
});

function animateButton(currentColor) {
    $("." + currentColor).addClass("pressed");
    setTimeout(function() {
        $("." + currentColor).removeClass("pressed");
    }, 100);
    var sound = new Audio("sounds/" + currentColor + ".mp3");
    sound.play();
}

function checkAnswer(currentLevel) {
    var success = 0;
    for(var i = 0; i < currentLevel; i++) {
        if(userClickedPattern[i] != gamePattern[i]) {
            console.log("wrong");
            $("#level-title").text("Game Over!");
            setTimeout(function() {
                location.reload();
            },5000);
        }
        else {
            console.log("success");
            success++;
        }
    }
    if(success === currentLevel){
        success = 0;
        userClickedPattern.splice(0, userClickedPattern.length);
        setTimeout(function() {
            nextSequence();
        },1000);
    }
}