function START() {
  var cg = new CodeGenerator();
  var code = "";
  
  // Headline
  code += cg.format("Auswertung - Runde " + cg.round).heading();
  code += cg.newParagraph();
  
  // Ranking of answers
  cg.questions.forEach(function (question, index) {
    var highestPoints = -1;
    
    code += cg.bbCode("u", (index + 1) + ". " + question);
    code += cg.newParagraph();
    cg.getDataOfAnswers(index + 1).forEach(function (data) {
      var points = data[0];
      var answer = data[1];
      var line = points + " " + answer;
      
      if (highestPoints === -1) {
        highestPoints = points;
      }
      
      code += highestPoints === points ? cg.bbCode("b", line) : line;
      code += cg.newLine();
    });
    code += cg.newLine();
  });
  code += cg.newParagraph();
  
  // Ranking of round
  var title = "Rangliste - Runde " + cg.round;
  var indexOfPointsColumn = _.INDEX_OF_POINTS_COLUMN;
  code += cg.ranking(title, indexOfPointsColumn);
  
  // Total ranking
  if (cg.round > 1) {
    title = "Gesamtwertung nach " + cg.round + " von 5 Runden";
    indexOfPointsColumn = _.INDEX_OF_TOTAL_COLUMN;
    code += cg.ranking(title, indexOfPointsColumn);
  }
  
  // Log the code
  Logger.log(cg.newParagraph() + code);
}


/**
* Constants
*/
var _ = {
  COLOR_OF_FIRST_POSITION: "#00FF00",
  COLOR_OF_SECOND_POSITION: "#00FFFF",
  COLOR_OF_THIRD_POSITION: "#FFFF00",
  HEADING_SIZE: 6,
  INDEX_OF_FIRST_ROW: 3,
  INDEX_OF_POINTS_COLUMN: 11,
  INDEX_OF_TOTAL_COLUMN: 13,
  INDEX_OF_USER_COLUMN: 0,
  QUESTIONS_RANGE: "P4:P8",
  ROUND_CELL: "C1"
};


/**
* CodeGenerator
*/
var CodeGenerator = function () {
  this.sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  this.rows = this.sheet.getDataRange().getValues().slice(_.INDEX_OF_FIRST_ROW);
  this.questions = this.sheet.getRange(_.QUESTIONS_RANGE).getValues();
  this.round = this.sheet.getRange(_.ROUND_CELL).getValue();
}

CodeGenerator.prototype.assignPositions = function (data) {
  var currentPosition = 0;
  var realPosition = 0;
  var lastPoints = data[0][0] + 1;
  
  return data.map(function (datum) {
    var points = datum[0];
    
    realPosition++;
    
    if (points < lastPoints) {
      lastPoints = points;
      currentPosition = realPosition;
    }
    
    datum.push(currentPosition);
    
    return datum;
  });
};

CodeGenerator.prototype.bbCode = function (tag, text) {
  return "[" + tag + "]" + text + "[/" + tag.split("=")[0] + "]";
};

CodeGenerator.prototype.format = function (text) {
  return {
    category: (function () {
      return this.bbCode("u", this.bbCode("b", text));
    }).bind(this),
    heading: (function () {
      return this.bbCode("size=" + _.HEADING_SIZE, this.format(text).category());
    }).bind(this),
    nthPosition: (function (nthPosition) {
      var colors = {
        1: _.COLOR_OF_FIRST_POSITION,
        2: _.COLOR_OF_SECOND_POSITION,
        3: _.COLOR_OF_THIRD_POSITION
      };
      
      if (nthPosition <= 3) {
        return this.bbCode("b", this.bbCode("color=" + colors[nthPosition], text));
      }
      
      return text;
    }).bind(this),
    points: (function (nthPosition) {
      if (nthPosition > 3) {
        text = this.bbCode("b", text);
      }
      
      return text;
    }).bind(this)
  };
};

CodeGenerator.prototype.getDataOfAnswers = function (nthRound) {
  var data = [];
  var seen = {};
  
  for (var i = 0; i < this.rows.length; i++) {
    var row = this.rows[i];
    var answerIndex = (nthRound - 1) * 2 + 1;
    var pointsIndex = (nthRound - 1) * 2 + 2;
    var answer = row[answerIndex];
    var points = row[pointsIndex];
    
    if (seen[answer] !== 1 && points !== 0) {
      seen[answer] = 1;
      data.push([points, answer]);
    }
  }
  
  return this.sortData(data);
};

CodeGenerator.prototype.getDataOfUsers = function (indexOfPointsColumn) {
  var data = this.rows.map(function (row) {
    var user = row[_.INDEX_OF_USER_COLUMN];
    var points = row[indexOfPointsColumn];
    
    return [points, user];
  }).filter(function (data) {
    return data[0] !== 0;
  });
  
  return this.sortData(data);
};

CodeGenerator.prototype.leadingZero = function (number) {
  return +number < 10 ? "0" + number : number;
};

CodeGenerator.prototype.newLine = function () {
  return "\n";
};

CodeGenerator.prototype.newParagraph = function () {
  return this.newLine() + this.newLine();
};

CodeGenerator.prototype.ranking = function (title, indexOfPointsColumn) {
  var code = "";
  
  code += this.format(title).category();
  code += this.newParagraph();
  this.assignPositions(this.getDataOfUsers(indexOfPointsColumn)).forEach((function (data) {
    var points = data[0];
    var user = data[1];
    var position = data[2];
    
    code += this.format(this.leadingZero(position) + ". " + this.format(points).points(position) + " " + user).nthPosition(position);
    code += this.newLine();
  }).bind(this));
  code += this.newParagraph();
  
  return code;
};

CodeGenerator.prototype.sortData = function (data) {
  return data.sort(function (a, b) {
    if (a[0] < b[0]) {
      return 1;
    }
    
    if (a[0] > b[0]) {
      return -1;
    }
    
    return 0;
  });
};
