/**
 * Implementation of @{link IEmmetEditor} interface
 * See `javascript/interfaces/IEmmetEditor.js`
 * @param {Function} require
 * @param {Underscore} _
 * @memberOf __Jacky.Emmet
 * @constructor
 */
emmet.define('editor', function (require, _) {

	return {
	    /** @memberOf editor */
	    createSelection: function (start, end) {
	        context.select(start, end);
	    },

	    getSelectionRange: function () {
	        var active = context.Active;
	        var anchor = context.Anchor;
	        if (active >= anchor) {
	            return {
	                start: anchor,
	                end: active
	            };
	        } else {
	            return {
	                start: active,
	                end: anchor
	            };
	        }
	    },

		getCurrentLineRange: function() {
			return {
				start: context.LineBegOffset,
				end: context.LineEndOffset
			};
		},

		getCaretPos: function() {
			return context.Active;
		},

		setCaretPos: function(pos) {
		    context.select(pos, pos);
		},

		getCurrentLine: function() {
			return context.Line;
		},

		replaceContent: function (value, start, end, noIndent) {
			if (_.isUndefined(end)) 
				end = _.isUndefined(start) ? content.length : start;
			if (_.isUndefined(start)) start = 0;
			var utils = require('utils');
		    
			// indent new value
			if (!noIndent) {
				value = utils.padString(value, utils.getLinePadding(this.getCurrentLine()));
			}
 
			// find new caret position
			var tabstopData = require('tabStops').extract(value, {
				escape: function(ch) {
					return ch;
				}
			});
			value = tabstopData.text;
			var firstTabStop = tabstopData.tabstops[0];
			
			if (firstTabStop) {
				firstTabStop.start += start;
				firstTabStop.end += start;
			} else {
				firstTabStop = {
					start: value.length + start,
					end: value.length + start
				};
			}
			context.replace(start, end, value);
			context.select(firstTabStop.start, firstTabStop.end);
		},

		getContent: function(){
		    return context.Text;
		},

		getSyntax: function () {
		    var path = context.Path;
		    var assocs = require('assocs').data();
		    for (var i = 0; i < assocs.length; i++) {
		        var assoc = assocs[i];
		        var pattern = new RegExp(assoc.pattern, "i");
		        if (pattern.test(path)) {
		            return assoc.syntax;
		        }
		    }
		    return context.Syntax;
		},

		getProfileName: function () {
		    var path = context.Path;
		    var assocs = require('assocs').data();
		    for (var i = 0; i < assocs.length; i++) {
		        var assoc = assocs[i];
		        var pattern = new RegExp(assoc.pattern, "i");
		        if (pattern.test(path)) {
		            return assoc.profile;
		        }
		    }
		    return context.Profile;
		},

		getSelection: function() {
		    return context.Selection;
		},

		getFilePath: function() {
			return context.Path;
		},

        prompt: function(title, value) {
            return context.prompt(title, value);
		}
	};
});
