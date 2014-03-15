/** 
 * Lorem Ipsum: Placeholder text generator add-on for Google Docs
 * Creator: Allen Guo
 * License: MIT
 * Website: https://github.com/guoguo12/lorem-ipsum
 */

function onOpen() {
  DocumentApp.getUi().createAddonMenu()
                     .addItem('Insert placeholder text...', 'showGenerateDialog')
                     .addToUi();
}

function onInstall() {
  onOpen();
}

function showGenerateDialog() {
  var props = PropertiesService.getDocumentProperties();
  var num = props.getProperty('last-used-num');
  var units = props.getProperty('last-used-units');
  var template = HtmlService.createTemplateFromFile('generate');
  template.num = num ? num : '2';
  template.units = units ? units : 'words';
  var dialog = template.evaluate();
  dialog.setHeight(80);
  dialog.setWidth(290);
  DocumentApp.getUi().showModelessDialog(dialog, 'Insert placeholder text');
}

function insertText(args) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty('last-used-num', args.count);
  props.setProperty('last-used-units', args.units);
  var gen = generator(args);
  DocumentApp.getActiveDocument().getCursor().insertText(gen);
}

/**
 * Everything below is taken directly from the LoremIpsum.js project
 * by Nickolas Kenyeres, which is freely available under the MIT License.
 *
 * Learn more at https://github.com/knicklabs/lorem-ipsum.js.
 */

var generator = function() {
  var options = (arguments.length) ? arguments[0] : {}
    , count = options.count || 1
    , units = options.units || 'sentences'
    , sentenceLowerBound = options.sentenceLowerBound || 5
    , sentenceUpperBound = options.sentenceUpperBound || 15
	  , paragraphLowerBound = options.paragraphLowerBound || 3
	  , paragraphUpperBound = options.paragraphUpperBound || 7
	  , format = options.format || 'plain'
    , words = options.words || dictionary.words
    , random = options.random || Math.random;

  units = simplePluralize(units.toLowerCase());

  var randomInteger = function(min, max) {
    return Math.floor(random() * (max - min + 1) + min);
  };
  
  var randomWord = function(words) {
    return words[randomInteger(0, words.length - 1)];
  };
  
  var randomSentence = function(words, lowerBound, upperBound) {
    var sentence = ''
      , bounds = {min: 0, max: randomInteger(lowerBound, upperBound)};
    
    while (bounds.min < bounds.max) {
      sentence = sentence + ' ' + randomWord(words);
      bounds.min = bounds.min + 1;
    }
    
    if (sentence.length) {
      sentence = sentence.slice(1);
      sentence = sentence.charAt(0).toUpperCase() + sentence.slice(1);
    }
  
    return sentence;
  };

  var randomParagraph = function(words, lowerBound, upperBound, sentenceLowerBound, sentenceUpperBound) {
    var paragraph = ''
      , bounds = {min: 0, max: randomInteger(lowerBound, upperBound)};
      
    while (bounds.min < bounds.max) {
      paragraph = paragraph + '. ' + randomSentence(words, sentenceLowerBound, sentenceUpperBound);
      bounds.min = bounds.min + 1;
    }
    
    if (paragraph.length) {
      paragraph = paragraph.slice(2);
      paragraph = paragraph + '.';
    }
    
    return paragraph;
  }
  
  var iter = 0
    , bounds = {min: 0, max: count}
    , string = ''
    , prefix = ''
    , suffix = "\r\n";

  if (format == 'html') {
    prefix = '<p>';
    suffix = '</p>';
  }
      
  while (bounds.min < bounds.max) {
    switch (units.toLowerCase()) {
      case 'words':
        string = string + ' ' + randomWord(words);
        break;
      case 'sentences':
        string = string + '. ' + randomSentence(words, sentenceLowerBound, sentenceUpperBound);
        break;
      case 'paragraphs':
        string = string + prefix + randomParagraph(words, paragraphLowerBound, paragraphUpperBound, sentenceLowerBound, sentenceUpperBound) + suffix;
        break;
    }
    bounds.min = bounds.min + 1;
  }
    
  if (string.length) {
    var pos = 0;
    
    if (string.indexOf('. ') == 0) {
      pos = 2;
    } else if (string.indexOf('.') == 0 || string.indexOf(' ') == 0) {
      pos = 1;
    }
    
    string = string.slice(pos);
    
    if (units == 'sentences') {
      string = string + '.';
    }
  }  
  
  return string;
};

function simplePluralize(string) {
  if (string.indexOf('s', string.length - 1) === -1) {
    return string + 's';
  }
  return string;
}

var dictionary = {
  words: [
    'ad',
    'adipisicing',
    'aliqua',
    'aliquip',
    'amet',
    'anim',
    'aute',
    'cillum',
    'commodo',
    'consectetur',
    'consequat',
    'culpa',
    'cupidatat',
    'deserunt',
    'do',
    'dolor',
    'dolore',
    'duis',
    'ea',
    'eiusmod',
    'elit',
    'enim',
    'esse',
    'est',
    'et',
    'eu',
    'ex',
    'excepteur',
    'exercitation',
    'fugiat',
    'id',
    'in',
    'incididunt',
    'ipsum',
    'irure',
    'labore',
    'laboris',
    'laborum',
    'Lorem',
    'magna',
    'minim',
    'mollit',
    'nisi',
    'non',
    'nostrud',
    'nulla',
    'occaecat',
    'officia',
    'pariatur',
    'proident',
    'qui',
    'quis',
    'reprehenderit',
    'sint',
    'sit',
    'sunt',
    'tempor',
    'ullamco',
    'ut',
    'velit',
    'veniam',
    'voluptate'  
  ]
};