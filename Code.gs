/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/*
    Copyright (C) 2017 IAmPicard
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Start STT', 'showSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('STT Crew loader');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function getPreferences() {
  var userProperties = PropertiesService.getUserProperties();
  return {accessToken: userProperties.getProperty('accessToken'), name: userProperties.getProperty('name')};
}

function login(username, password) {
  var data = 'username=' + username + '&password=' + password + '&client_id=4fc852d7-d602-476a-a292-d243022a475d&grant_type=password';

  var options = {
   'method' : 'post',
   'payload' : data
  };

  var response = UrlFetchApp.fetch('https://thorium.disruptorbeam.com/oauth2/token', options);
  return JSON.parse(response.getContentText());
}

function loadCrew(access_token) {
  var response = UrlFetchApp.fetch('https://stt.disruptorbeam.com/player?client_api=8&access_token=' + access_token);
  var playerData = JSON.parse(response.getContentText());

  response = UrlFetchApp.fetch('https://stt.disruptorbeam.com/character/get_avatar_crew_archetypes?client_api=8&access_token=' + access_token);
  var crewArchetypes = JSON.parse(response.getContentText());

  var result = {
    vipLevel: playerData.player.vip_level,
    name: playerData.player.character.display_name,
    level: playerData.player.character.level,
    crewLimit: playerData.player.character.crew_limit,
    crew: undefined,
    cadetMissions: []
  };

  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('accessToken', access_token);
  userProperties.setProperty('name', result.name);

  result.crew = new Object();
  crewArchetypes.crew_avatars.forEach(function (av) {
    result.crew[av.id] = {name: av.name, short_name: av.short_name, max_rarity: av.max_rarity, have: false, airlocked: false, immortal: 0};
  });

  crewArchetypes = undefined;

  playerData.player.character.crew.forEach(function (crew) {
    result.crew[crew.archetype_id].have = true;
    result.crew[crew.archetype_id].flavor = crew.flavor;
    result.crew[crew.archetype_id].level = crew.level;
    result.crew[crew.archetype_id].rarity = crew.rarity;
    result.crew[crew.archetype_id].traits = crew.traits;
    result.crew[crew.archetype_id].traits_hidden = crew.traits_hidden;
    result.crew[crew.archetype_id].skills = crew.skills;
    result.crew[crew.archetype_id].ship_battle = crew.ship_battle;
    result.crew[crew.archetype_id].equipment = crew.equipment.length;
    result.crew[crew.archetype_id].favorite = crew.favorite;
    result.crew[crew.archetype_id].airlocked = crew.in_buy_back_state;
  });

  playerData.player.character.stored_immortals.forEach(function (imm) {
    result.crew[imm.id].immortal = imm.quantity;
  });

  playerData.player.character.cadet_schedule.missions.forEach(function (mission) {
    result.cadetMissions.push(mission.id);
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Crew roster");
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet('Crew roster');

  sheet.appendRow([' ', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
  sheet.appendRow(["Name", "Have", "Rarity", "Max", "Level", "Equipment", 'Core', 'Min', 'Max', 'Core', 'Min', 'Max', 'Core', 'Min', 'Max', 'Core', 'Min', 'Max', 'Core', 'Min', 'Max', 'Core', 'Min', 'Max', "Traits"]);

  // TODO: Can this be shared between the 2 functions?
  var SKILLS = {
    'command_skill': 'Command',
    'science_skill': 'Science',
    'security_skill': 'Security',
    'engineering_skill': 'Engineering',
    'diplomacy_skill': 'Diplomacy',
    'medicine_skill': 'Medicine'
  };

  var RARITYCOLORS = [
    { b : '', f: '' }, // Basic
    { b : '#9b9b9b', f: '#080808' }, // Common
    { b : '#50aa3c', f: '#963caa' }, // Uncommon
    { b : '#5aaaff', f: '#ffaf5a' }, // Rare
    { b : '#aa2deb', f: '#6eeb2d' }, // Super Rare
    { b : '#fdd26a', f: '#6a95fd' } // Legendary
  ];

  var colIndex = 7;
  for (var skill in SKILLS) {
    var crew = SKILLS[skill];

    var range = sheet.getRange(1, colIndex, 1, 3);
    range.merge();
    range.setValue(SKILLS[skill]);
    range.setFontWeight("bold");
    range.setHorizontalAlignment("center");

    colIndex = colIndex + 3;
  }

  for (var crewId in result.crew) {
    var crew = result.crew[crewId];

    if (crew.have) {
      sheet.appendRow([
        crew.name,
        'Yes',
        crew.rarity,
        crew.max_rarity,
        crew.level,
        '' + crew.equipment + ' / 4',
        crew.skills.command_skill ? crew.skills.command_skill.core : 0,
        crew.skills.command_skill ? crew.skills.command_skill.range_min : 0,
        crew.skills.command_skill ? crew.skills.command_skill.range_max : 0,
        crew.skills.science_skill ? crew.skills.science_skill.core : 0,
        crew.skills.science_skill ? crew.skills.science_skill.range_min : 0,
        crew.skills.science_skill ? crew.skills.science_skill.range_max : 0,
        crew.skills.security_skill ? crew.skills.security_skill.core : 0,
        crew.skills.security_skill ? crew.skills.security_skill.range_min : 0,
        crew.skills.security_skill ? crew.skills.security_skill.range_max : 0,
        crew.skills.engineering_skill ? crew.skills.engineering_skill.core : 0,
        crew.skills.engineering_skill ? crew.skills.engineering_skill.range_min : 0,
        crew.skills.engineering_skill ? crew.skills.engineering_skill.range_max : 0,
        crew.skills.diplomacy_skill ? crew.skills.diplomacy_skill.core : 0,
        crew.skills.diplomacy_skill ? crew.skills.diplomacy_skill.range_min : 0,
        crew.skills.diplomacy_skill ? crew.skills.diplomacy_skill.range_max : 0,
        crew.skills.medicine_skill ? crew.skills.medicine_skill.core : 0,
        crew.skills.medicine_skill ? crew.skills.medicine_skill.range_min : 0,
        crew.skills.medicine_skill ? crew.skills.medicine_skill.range_max : 0,
        crew.traits.join(', ') + ', ' + crew.traits_hidden.join(', ')
      ]);
    } else {
      if (crew.immortal > 0) {
        sheet.appendRow([
          crew.name,
          'Vault',
          crew.max_rarity,
          crew.max_rarity
        ]);
      } else {
         sheet.appendRow([
            crew.name,
            'No',
            0,
            crew.max_rarity
          ]);
      }
    }
    
    var rangeRarity = sheet.getRange(sheet.getLastRow(), 3, 1, 2);
    rangeRarity.setBackground(RARITYCOLORS[crew.max_rarity].b);
    //rangeRarity.setFontColor(RARITYCOLORS[crew.max_rarity].f);
  }

  // Freeze the first 2 rows
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);

  for (var i = 1; i < sheet.getLastColumn(); i++) {
    sheet.autoResizeColumn(i);
  }

  for (var i = 7; i < 25; i++) {
    sheet.setColumnWidth(i, 38);
  }

  colIndex = 7;
  for (var i = 0; i < 7; i++) {
    var range = sheet.getRange(1, colIndex, sheet.getLastRow());
    range.setBorder(null, true, null, null, false, false, null, null); // SpreadsheetApp.BorderStyle.SOLID_THICK);
    colIndex = colIndex + 3;
  }

  sheet.showSheet();

  return result;
}

function loadCadetMissionData(cadetMissions, access_token) {
  var missionIds = '';
  cadetMissions.forEach(function (missionId) {
    missionIds = missionIds + 'ids[]=' + missionId + '&';
  });

  var response = UrlFetchApp.fetch('https://stt.disruptorbeam.com/mission/info?' + missionIds + 'client_api=8&access_token=' + access_token);
  var missionData = JSON.parse(response.getContentText());

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Cadet missions");
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet('Cadet missions');

  sheet.appendRow(["Mission", "Conflict", "Challenge", "Skill", "Difficulty", 'Crit', 'Traits', 'Trait bonus', 'Min Stars', 'Max Stars', 'Required Traits']);

  var SKILLS = {
    'command_skill': 'Command',
    'science_skill': 'Science',
    'security_skill': 'Security',
    'engineering_skill': 'Engineering',
    'diplomacy_skill': 'Diplomacy',
    'medicine_skill': 'Medicine'
  };

  function getCrit(challenge) {
    if (!challenge.critical) {
      return 'None';
    }

    if (challenge.critical.claimed == true) {
      return 'Claimed (' + challenge.critical.threshold + ')';
    }

    return 'Unclaimed (' + challenge.critical.threshold + ')';
  }

  missionData.character.accepted_missions.forEach(function (mission) {
    mission.quests.forEach(function (quest) {
      if (quest.quest_type == 'ConflictQuest') {
        response = UrlFetchApp.fetch('https://stt.disruptorbeam.com/quest/conflict_info?id=' + quest.id + '&client_api=8&access_token=' + access_token);
        var questData = JSON.parse(response.getContentText());
        questData.challenges.forEach(function (challenge) {
          var traits = [];
          var bonus = 0;
          challenge.trait_bonuses.forEach(function (traitBonus) {
            bonus = traitBonus.bonuses[2];
            traits.push(traitBonus.trait);
          });

          sheet.appendRow([
            mission.episode_title,
            questData.name,
            challenge.name,
            SKILLS[challenge.skill],
            challenge.difficulty_by_mastery[2],
            getCrit(challenge),
            traits.join(', '),
            bonus,
            questData.crew_requirement.min_stars,
            questData.crew_requirement.max_stars,
            questData.crew_requirement.traits ? questData.crew_requirement.traits.join(', ') : ''
          ]);

          if (challenge.critical && challenge.critical.claimed == false) {
            var rangeCrit = sheet.getRange(sheet.getLastRow(), 6);
            rangeCrit.setBackground('red');
          }
        });
      }
    });
  });

  // Freeze the first row
  sheet.setFrozenRows(1);

  for (var i = 1; i < sheet.getLastColumn(); i++) {
    sheet.autoResizeColumn(i);
  }

  sheet.showSheet();
}
