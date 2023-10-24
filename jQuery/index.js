$("h1").addClass("big-title margin-50");

$("h1").text("Good Bye");

$("a").attr("href", "https://www.google.com");

$("h1").click(function() {
    $("h1").css("color", "purple");
});

$("h1").on("mouseover", function() {
    $("h1").css("font-size", "2rem");
});

$("input").keydown(function(event) {
    $("h1").text(event.key);
});

$(".hide").on("click", function() {
    $("h1").hide(); //.fadeOut() .slideUp()
});

$(".show").on("click", function() {
    $("h1").show(); //.fadeIn() .slideDown()
});

$(".toggle").on("click", function() {
    $("h1").toggle(); //.fadeToggle() .slideToggle()
});

$(".animate").on("click", function() {
    $("h1").animate({opacity: 0}).animate({opacity: 1}); //css that uses numberic value
});
