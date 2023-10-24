/* Creating an object */

var houseKeeper1 = {
    yearsOfExperience: 12,
    name: "Jane",
    cleaningRepertoire: ["bathroom", "lobby", "bedroom"]
}

/* Constructor Function */
/* construction function should be name with the first letter capitalized */

function HouseKeeper(yearsOfExperience, name, cleaningRepertoire) {
    this.yearsOfExperience = yearsOfExperience;
    this.name = name;
    this.cleaningRepertoire = cleaningRepertoire
    this.clean = function() {
        alert("Cleaning in progress.");
    }
}

var houseKeeper1 = new HouseKeeper(12, "Jane", ["bathroom", "lobby", "bedroom"]);

houseKeeper1.clean();