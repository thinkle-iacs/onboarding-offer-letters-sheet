/**
 * Generates a relatively secure initial password.
 * @param {number} length - The desired length of the password.
 * @return {string} - The generated password.
 */ /**
 * Generates a relatively secure and readable initial password.
 * @param {number} length - The desired length of the password.
 * @return {string} - The generated password.
 */
function generatePassword(length = 12) {
  const stems = [
    "teach",
    "learn",
    "rock",
    "star",
    "go",
    "run",
    "sun",
    "fun",
    "math",
    "read",
    "write",
    "lab",
    "code",
    "bike",
    "swim",
    "jump",
    "play",
    "quiz",
    "solve",
    "build",
    "think",
    "grow",
    "dream",
    "plan",
    "craft",
    "draw",
    "sketch",
    "plan",
    "play",
    "guide",
    "paint",
    "study",
    "field",
    "woods",
    "track",
    "coach",
    "walk",
    "talk",
    "game",
    "jump",
    "hike",
    "bike",
    "run",
    "climb",
    "trip",
    "camp",
    "wave",
    "surf",
    "dive",
    "ride",
    "read",
    "play",
    "quiz",
    "test",
    "work",
    "rest",
    "plan",
    "help",
    "pear",
    "peach",
    "road",
    "path",
    "moon",
    "key",
    "see",
    "hear",
    "feel",
    "cpu",
    "clock",
    "reach",
    "teach",
    "learn",
    "rock",
    "star",
    "go",
    "run",
    "sun",
    "fun",
    "leaf",
    "tree",
    "bush",
    "bloom",
    "blossom",
    "rose",
    "lily",
    "daisy",
    "tulip",
    "iris",
    "daffodil",
    "crocus",
    "violet",
    "poppy",
    "peony",
    "dahlia",
    "aster",
    "beach",
    "sea",
    "canoe",
    "kayak",
    "raft",
    "sail",
    "surf",
    "dive",
    "fish",
    "swim",
    "splash",
    "jump",
    "run",
    "walk",
    "hike",
    "bike",
    "climb",
    "ski",
    "skate",
  ];

  const symbols = "123456789!@#$%*=+?";

  let password = "";
  while (password.length < length) {
    // Add a random stem
    const stemIndex = Math.floor(Math.random() * stems.length);
    password += stems[stemIndex];

    // Add two  random symbols

    password += symbols[Math.floor(Math.random() * symbols.length)];
    password += symbols[Math.floor(Math.random() * symbols.length)];
  }

  return password;
}
function testPass() {
  console.log(generatePassword(12));
  console.log(generatePassword(20));
}
