/**
 * @param {Function} require
 * @param {Underscore} _
 * @memberOf __Jacky.Emmet
 * @constructor
 */
emmet.define('assocs', function (require, _) {
    var settings = JSON.parse(require('file').read(context.Root + '\\Assets\\settings.json'));
    return {
        data: function() {
            return settings;
        }
    }
});
