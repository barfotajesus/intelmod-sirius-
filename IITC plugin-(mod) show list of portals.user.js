// ==UserScript==
// @id             iitc-plugin-portals-list@teo96
// @name           IITC plugin:(mod) show list of portals
// @category       Info
// @version        0.2.1.20150917.154202
// @namespace      https://github.com/jonatkins/ingress-intel-total-conversion
// @updateURL      https://secure.jonatkins.com/iitc/release/plugins/portals-list.meta.js
// @downloadURL    https://secure.jonatkins.com/iitc/release/plugins/portals-list.user.js
// @description    [jonatkins-2015-09-17-154202] Display a sortable list of all visible portals with full details about the team, resonators, links, etc.
// @include        https://www.ingress.com/intel*
// @include        http://www.ingress.com/intel*
// @match          https://www.ingress.com/intel*
// @match          http://www.ingress.com/intel*
// @include        https://www.ingress.com/mission/*
// @include        http://www.ingress.com/mission/*
// @match          https://www.ingress.com/mission/*
// @match          http://www.ingress.com/mission/*
// @grant          none
// ==/UserScript==


function wrapper(plugin_info) {
// ensure plugin framework is there, even if iitc is not yet loaded
if(typeof window.plugin !== 'function') window.plugin = function() {};

//PLUGIN AUTHORS: writing a plugin outside of the IITC build environment? if so, delete these lines!!
//(leaving them in place might break the 'About IITC' page or break update checks)
plugin_info.buildName = 'jonatkins';
plugin_info.dateTimeVersion = '20150917.154202';
plugin_info.pluginId = 'portals-list';
//END PLUGIN AUTHORS NOTE



// PLUGIN START ////////////////////////////////////////////////////////

// use own namespace for plugin
window.plugin.portalslistmod = function() {};

window.plugin.portalslistmod.listPortals = [];
window.plugin.portalslistmod.sortBy = 1; // second column: level
window.plugin.portalslistmod.sortOrder = -1;
window.plugin.portalslistmod.enlP = 0;
window.plugin.portalslistmod.resP = 0;
window.plugin.portalslistmod.neuP = 0;
window.plugin.portalslistmod.filter = 0;

/*
 * plugins may add fields by appending their specifiation to the following list. The following members are supported:
 * title: String
 *     Name of the column. Required.
 * value: function(portal)
 *     The raw value of this field. Can by anything. Required, but can be dummy implementation if sortValue and format
 *     are implemented.
 * sortValue: function(value, portal)
 *     The value to sort by. Optional, uses value if omitted. The raw value is passed as first argument.
 * sort: function(valueA, valueB, portalA, portalB)
 *     Custom sorting function. See Array.sort() for details on return value. Both the raw values and the portal objects
 *     are passed as arguments. Optional. Set to null to disable sorting
 * format: function(cell, portal, value)
 *     Used to fill and format the cell, which is given as a DOM node. If omitted, the raw value is put in the cell.
 * defaultOrder: -1|1
 *     Which order should by default be used for this column. -1 means descending. Default: 1
 */


window.plugin.portalslistmod.fields = [
  {
    title: "Portal Name",
    value: function(portal) { return portal.options.data.title; },
    sortValue: function(value, portal) { return value.toLowerCase(); },
    format: function(cell, portal, value) {
      $(cell)
        .append(plugin.portalslistmod.getPortalLink(portal))
        .addClass("portalTitle");
    }
  },
  {
    title: "Mod",
      value: function(portal) { return 'tmod' in portal.options.data ? portal.options.data.tmod : '-'; },
    format: function(cell, portal, value) {
      $(cell)
        .text(value);
    },
    defaultOrder: -1,
  },
  {
    title: "Lv",
    value: function(portal) { return portal.options.data.level; },
    format: function(cell, portal, value) {
      $(cell)
        .css('background-color', COLORS_LVL[value])
        .text('L' + value);
    },
    defaultOrder: -1,
  },
  {
      title: "Tm",
    value: function(portal) { return portal.options.team; },
    format: function(cell, portal, value) {
      $(cell).text(['NEU', 'RES', 'ENL'][value]);
    }
  },
  {
    title: "O",
    value: function(portal) { return 'tow' in portal.options.data ? portal.options.data.tow : '0'; },
    format: function(cell, portal, value) {
      $(cell)
        .text([' ', 'R', 'O'][value]);
    },
    defaultOrder: -1,
  },
/*  { 
    title: "Owner", 
    value: function(portal) { return 'town' in portal.options.data ? portal.options.data.town : '-' }, 
    sortValue: function(value, portal) { return 'town' in portal.options.data ? portal.options.data.town : '0' }, 
    format: function(cell, portal, value) { 
      $(cell) 
        .text(value); 
    }, 
    defaultOrder: -1, 
  }, */
  {
    title: "@@",
    value: function(portal) { return 'need' in portal.options.data ? portal.options.data.need : '-'; },
    sortValue: function(value, portal) { return 'need' in portal.options.data ? (9 - portal.options.data.need) : '0'; },
    format: function(cell, portal, value) {
      $(cell)
        .addClass("alignR")
        .text(portal.options.team===TEAM_NONE ? '8' : value);
    },
    defaultOrder: -1,
  },
  {
    title: "Hack",
    value: function(portal) { return 'th1' in portal.options.data ? (portal.options.data.th1 + "[" + portal.options.data.th2 +"s]") : '-'; },
    sortValue: function(value, portal) { return 'th2' in portal.options.data ? portal.options.data.th2 : '400'; },
    format: function(cell, portal, value) {
      $(cell)
        .text(value);
    },
    defaultOrder: -1,
  },
  {
    title: "Sheld",
    value: function(portal) { return 'tg1' in portal.options.data ? (portal.options.data.tg1 + "(+" + portal.options.data.tg2 +")") : '-'; },
    sortValue: function(value, portal) { return 'tg1' in portal.options.data ? (portal.options.data.tg1 + portal.options.data.tg2) : '0'; },
    format: function(cell, portal, value) {
      $(cell)
        .text(value);
    },
    defaultOrder: -1,
  },
  {
    title: "Lat",
    value: function(portal) { return portal.options.data.latE6; },
    format: function(cell, portal, value) {
      $(cell)
        .text(value);
    },
    defaultOrder: -1,
  },
  {
    title: "Lng",
    value: function(portal) { return portal.options.data.lngE6; },
    format: function(cell, portal, value) {
      $(cell)
        .text(value);
    },
    defaultOrder: -1,
  },
  {
    title: "URL",
    value: function(portal) { return plugin.portalslistmod.getPortalLink(portal); },
    format: function(cell, portal, value) {
      $(cell)
        .text(value);
    },
    defaultOrder: -1,
  },
];
    
//fill the listPortals array with portals avaliable on the map (level filtered portals will not appear in the table)
window.plugin.portalslistmod.getPortals = function() {
  //filter : 0 = All, 1 = Neutral, 2 = Res, 3 = Enl, -x = all but x
  var retval=false;

  var displayBounds = map.getBounds();

  window.plugin.portalslistmod.listPortals = [];
  $.each(window.portals, function(i, portal) {
    // eliminate offscreen portals (selected, and in padding)
    if(!displayBounds.contains(portal.getLatLng())) return true;

    retval=true;

    switch (portal.options.team) {
      case TEAM_RES:
        window.plugin.portalslistmod.resP++;
        break;
      case TEAM_ENL:
        window.plugin.portalslistmod.enlP++;
        break;
      default:
        window.plugin.portalslistmod.neuP++;
    }

    // cache values and DOM nodes
    var obj = { portal: portal, values: [], sortValues: [] };

    var row = document.createElement('tr');
    row.className = TEAM_TO_CSS[portal.options.team];
    obj.row = row;

    var cell = row.insertCell(-1);
    cell.className = 'alignR';

    window.plugin.portalslistmod.fields.forEach(function(field, i) {
      cell = row.insertCell(-1);

      var value = field.value(portal);
      obj.values.push(value);

      obj.sortValues.push(field.sortValue ? field.sortValue(value, portal) : value);

      if(field.format) {
        field.format(cell, portal, value);
      } else {
        cell.textContent = value;
      }
    });

    window.plugin.portalslistmod.listPortals.push(obj);
  });

  return retval;
};

window.plugin.portalslistmod.displayPL = function() {
  var list;
  // plugins (e.g. bookmarks) can insert fields before the standard ones - so we need to search for the 'level' column
  window.plugin.portalslistmod.sortBy = window.plugin.portalslistmod.fields.map(function(f){return f.title;}).indexOf('Lv');
  window.plugin.portalslistmod.sortOrder = -1;
  window.plugin.portalslistmod.enlP = 0;
  window.plugin.portalslistmod.resP = 0;
  window.plugin.portalslistmod.neuP = 0;
  window.plugin.portalslistmod.filter = 0;

  if (window.plugin.portalslistmod.getPortals()) {
    list = window.plugin.portalslistmod.portalTable(window.plugin.portalslistmod.sortBy, window.plugin.portalslistmod.sortOrder,window.plugin.portalslistmod.filter);
  } else {
    list = $('<table class="noPortals"><tr><td>Nothing to show!</td></tr></table>');
  }

  if(window.useAndroidPanes()) {
    $('<div id="portalslistmod" class="mobile">').append(list).appendTo(document.body);
  } else {
    dialog({
      html: $('<div id="portalslistmod">').append(list),
      dialogClass: 'ui-dialog-portalslistmod',
      title: 'Portal list: ' + window.plugin.portalslistmod.listPortals.length + ' ' + (window.plugin.portalslistmod.listPortals.length == 1 ? 'portal' : 'portals'),
      id: 'portal-list',
      width: 900
    });
  }
};

window.plugin.portalslistmod.portalTable = function(sortBy, sortOrder, filter) {
  // save the sortBy/sortOrder/filter
  window.plugin.portalslistmod.sortBy = sortBy;
  window.plugin.portalslistmod.sortOrder = sortOrder;
  window.plugin.portalslistmod.filter = filter;

  var portals = window.plugin.portalslistmod.listPortals;
  var sortField = window.plugin.portalslistmod.fields[sortBy];

  portals.sort(function(a, b) {
    var valueA = a.sortValues[sortBy];
    var valueB = b.sortValues[sortBy];

    if(sortField.sort) {
      return sortOrder * sortField.sort(valueA, valueB, a.portal, b.portal);
    }

//FIXME: sort isn't stable, so re-sorting identical values can change the order of the list.
//fall back to something constant (e.g. portal name?, portal GUID?),
//or switch to a stable sort so order of equal items doesn't change
    return sortOrder *
      (valueA < valueB ? -1 :
      valueA > valueB ?  1 :
      0);
  });

  if(filter !== 0) {
    portals = portals.filter(function(obj) {
      return filter < 0
        ? obj.portal.options.team+1 != -filter
        : obj.portal.options.team+1 == filter;
    });
  }

  var table, row, cell;
  var container = $('<div>');

  table = document.createElement('table');
  table.className = 'filter';
  container.append(table);

  row = table.insertRow(-1);

  var length = window.plugin.portalslistmod.listPortals.length;

  ["All", "Neutral", "Resistance", "Enlightened"].forEach(function(label, i) {
    cell = row.appendChild(document.createElement('th'));
    cell.className = 'filter' + label.substr(0, 3);
    cell.textContent = label+':';
    cell.title = 'Show only portals of this color';
    $(cell).click(function() {
      $('#portalslistmod').empty().append(window.plugin.portalslistmod.portalTable(sortBy, sortOrder, i));
    });


    cell = row.insertCell(-1);
    cell.className = 'filter' + label.substr(0, 3);
    if(i != 0) cell.title = 'Hide portals of this color';
    $(cell).click(function() {
      $('#portalslistmod').empty().append(window.plugin.portalslistmod.portalTable(sortBy, sortOrder, -i));
    });
    switch(i-1) {
      case -1:
        cell.textContent = length;
        break;
      case 0:
        cell.textContent = window.plugin.portalslistmod.neuP + ' (' + Math.round(window.plugin.portalslistmod.neuP/length*100) + '%)';
        break;
      case 1:
        cell.textContent = window.plugin.portalslistmod.resP + ' (' + Math.round(window.plugin.portalslistmod.resP/length*100) + '%)';
        break;
      case 2:
        cell.textContent = window.plugin.portalslistmod.enlP + ' (' + Math.round(window.plugin.portalslistmod.enlP/length*100) + '%)';
    }
  });

  table = document.createElement('table');
  table.className = 'portals';
  container.append(table);

  var thead = table.appendChild(document.createElement('thead'));
  row = thead.insertRow(-1);

  cell = row.appendChild(document.createElement('th'));
  cell.textContent = '#';

  window.plugin.portalslistmod.fields.forEach(function(field, i) {
    cell = row.appendChild(document.createElement('th'));
    cell.textContent = field.title;
    if(field.sort !== null) {
      cell.classList.add("sortable");
      if(i == window.plugin.portalslistmod.sortBy) {
        cell.classList.add("sorted");
      }

      $(cell).click(function() {
        var order;
        if(i == sortBy) {
          order = -sortOrder;
        } else {
          order = field.defaultOrder < 0 ? -1 : 1;
        }

        $('#portalslistmod').empty().append(window.plugin.portalslistmod.portalTable(i, order, filter));
      });
    }
  });

  portals.forEach(function(obj, i) {
    var row = obj.row
    if(row.parentNode) row.parentNode.removeChild(row);

    row.cells[0].textContent = i+1;

    table.appendChild(row);
  });
//@@ ここから変更箇所
  var t_txt ='';
  var ty,tx;
  for(ty in portals) {
//    if(6 < portals[ty].portal.options.level || portals[ty].values[1].match(/VRH/) || portals[ty].values[1].match(/VRM/)){
//    portals[ty].values[5]... 表示されている値の5番目(名前が0)
//    portals[ty].values[1].match(/RM/)... RMという文字があったら
//    if( portals[ty].values[5] < 4 || portals[ty].values[1].match(/RH/) || portals[ty].values[1].match(/RM/)){
    if( ((portals[ty].values[5] < 4) & (portals[ty].values[3] ==1))
       || ((portals[ty].values[5] < 1 & portals[ty].values[3] ==2))
       || portals[ty].values[1].match(/RH/)
       || portals[ty].values[1].match(/RM/)
      ){
      for(tx in portals[ty].values) {
        t_txt = t_txt + portals[ty].values[tx] + '\t';
      }
      t_txt = t_txt + '\n';
    }
  }
var data = t_txt;
var a = document.createElement('a');
a.textContent = 'export';
a.download = 'table.csv';
a.href = window.URL.createObjectURL(new Blob([data], { type: 'text/plain' }));
a.dataset.downloadurl = ['text/plain', a.download, a.href].join(':');
 
  container.append('<div class="disclaimer">Click on portals table headers to sort by that column. '
    + 'Click on <b>All, Neutral, Resistance, Enlightened</b> to only show portals owner by that faction or on the number behind the factions to show all but those portals.'
    + ' [<a href=' + a.href + ' download=test.csv ><span id="export-link">export</span></a>]</div>');
//元の文言
//  container.append('<div class="disclaimer">Click on portals table headers to sort by that column. '
//    + 'Click on <b>All, Neutral, Resistance, Enlightened</b> to only show portals owner by that faction or on the number behind the factions to show all but those portals.</div>');
//@@

  return container;
};

// portal link - single click: select portal
//               double click: zoom to and select portal
// code from getPortalLink function by xelio from iitc: AP List - https://raw.github.com/breunigs/ingress-intel-total-conversion/gh-pages/plugins/ap-list.user.js
window.plugin.portalslistmod.getPortalLink = function(portal) {
  var coord = portal.getLatLng();
  var perma = '/intel?ll='+coord.lat+','+coord.lng+'&z=17&pll='+coord.lat+','+coord.lng;

  // jQuery's event handlers seem to be removed when the nodes are remove from the DOM
  var link = document.createElement("a");
  link.textContent = portal.options.data.title;
  link.href = perma;
  link.addEventListener("click", function(ev) {
    renderPortalDetails(portal.options.guid);
    ev.preventDefault();
    return false;
  }, false);
  link.addEventListener("dblclick", function(ev) {
//    zoomToAndShowPortal(portal.options.guid, [coord.lat, coord.lng]);
    ev.preventDefault();
    return false;
  });
  return link;
};

window.plugin.portalslistmod.onPaneChanged = function(pane) {
  if(pane == "plugin-portalslistmod")
    window.plugin.portalslistmod.displayPL();
  else
    $("#portalslistmod").remove();
};

var setup =  function() {
  if(window.useAndroidPanes()) {
    android.addPane("plugin-portalslistmod", "Portals list(mod)", "ic_action_paste");
    addHook("paneChanged", window.plugin.portalslistmod.onPaneChanged);
  } else {
    $('#toolbox').append('<a onclick="window.plugin.portalslistmod.displayPL()" title="Display a list of portals in the current view [t]" accesskey="t">Portals(mod)</a>');
  }

  $("<style>")
    .prop("type", "text/css")
    .html("#portalslistmod.mobile {\n  background: transparent;\n  border: 0 none !important;\n  height: 100% !important;\n  width: 100% !important;\n  left: 0 !important;\n  top: 0 !important;\n  position: absolute;\n  overflow: auto;\n}\n\n#portalslistmod table {\n  margin-top: 5px;\n  border-collapse: collapse;\n  empty-cells: show;\n  width: 100%;\n  clear: both;\n}\n\n#portalslistmod table td, #portalslistmod table th {\n  background-color: #1b415e;\n  border-bottom: 1px solid #0b314e;\n  color: white;\n  padding: 3px;\n}\n\n#portalslistmod table th {\n  text-align: center;\n}\n\n#portalslistmod table .alignR {\n  text-align: right;\n}\n\n#portalslistmod table.portals td {\n  white-space: nowrap;\n}\n\n#portalslistmod table th.sortable {\n  cursor: pointer;\n}\n\n#portalslistmod table .portalTitle {\n  min-width: 120px !important;\n  max-width: 240px !important;\n  overflow: visible;\n  white-space: nowrap;\n  text-overflow: ellipsis;\n}\n\n#portalslistmod .sorted {\n  color: #FFCE00;\n}\n\n#portalslistmod table.filter {\n  table-layout: fixed;\n  cursor: pointer;\n  border-collapse: separate;\n  border-spacing: 1px;\n}\n\n#portalslistmod table.filter th {\n  text-align: left;\n  padding-left: 0.3em;\n  overflow: hidden;\n  text-overflow: ellipsis;\n}\n\n#portalslistmod table.filter td {\n  text-align: right;\n  padding-right: 0.3em;\n  overflow: hidden;\n  text-overflow: ellipsis;\n}\n\n#portalslistmod .filterNeu {\n  background-color: #666;\n}\n\n#portalslistmod table tr.res td, #portalslistmod .filterRes {\n  background-color: #005684;\n}\n\n#portalslistmod table tr.enl td, #portalslistmod .filterEnl {\n  background-color: #017f01;\n}\n\n#portalslistmod table tr.none td {\n  background-color: #000;\n}\n\n#portalslistmod .disclaimer {\n  margin-top: 10px;\n  font-size: 10px;\n}\n\n#portalslistmod.mobile table.filter tr {\n  display: block;\n  text-align: center;\n}\n#portalslistmod.mobile table.filter th, #portalslistmod.mobile table.filter td {\n  display: inline-block;\n  width: 22%;\n}\n\n")
    .appendTo("head");

};

// PLUGIN END //////////////////////////////////////////////////////////


setup.info = plugin_info; //add the script info data to the function as a property
if(!window.bootPlugins) window.bootPlugins = [];
window.bootPlugins.push(setup);
// if IITC has already booted, immediately run the 'setup' function
if(window.iitcLoaded && typeof setup === 'function') setup();
} // wrapper end
// inject code into site context
var script = document.createElement('script');
var info = {};
if (typeof GM_info !== 'undefined' && GM_info && GM_info.script) info.script = { version: GM_info.script.version, name: GM_info.script.name, description: GM_info.script.description };
script.appendChild(document.createTextNode('('+ wrapper +')('+JSON.stringify(info)+');'));
(document.body || document.head || document.documentElement).appendChild(script);
