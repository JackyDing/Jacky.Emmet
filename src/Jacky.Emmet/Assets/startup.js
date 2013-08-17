/**
 * import underscore module
 * import core module
 */
context.require(context.Root + '\\Assets\\javascript\\underscore.js');
context.require(context.Root + '\\Assets\\javascript\\json.js');
context.require(context.Root + '\\Assets\\javascript\\core.js');
context.require(context.Root + '\\Assets\\javascript\\file.js');

/**
 * Set module loader and construct editor instance
 *
 * @param {Function} require
 * @param {Underscore} _
 * @memberOf __Jacky.Emmet
 */
var editor = emmet.exec(function(require, _) {
    
    function getModuleLoader() {
        var paths = _.toArray(arguments);
        return function (module) {
            _.find(paths, function (path) {
                var url = path + '\\' + module + '.js';
                return context.require(url);
            });
        };
    }

	emmet.setModuleLoader(getModuleLoader(
        context.Root + '\\Assets',
        context.Root + '\\Assets\\javascript',
        context.Root + '\\Assets\\javascript\\loaders',
        context.Root + '\\Assets\\javascript\\parsers',
        context.Root + '\\Assets\\javascript\\parsers\\editTree',
        context.Root + '\\Assets\\javascript\\resolvers'
    ));

	return require('editor');
});

/**
 * Load system snippets, load buildin actions
 *
 * @param {Function} require
 * @param {Underscore} _
 * @memberOf __Jacky.Emmet
 */
emmet.exec(function (require, _) {
    require('filters\\bem');
    require('filters\\comment');
    require('filters\\escape');
    require('filters\\format');
    require('filters\\haml');
    require('filters\\html');
    require('filters\\single-line');
    require('filters\\trim');
    require('filters\\xsl');
    require('processors\\tag-name');
    require('processors\\pasted-content');
    require('processors\\resource-matcher');
    require('generators\\lorem-ipsum');
    require('actions\\expandAbbreviation');
    require('actions\\wrapWithAbbreviation');
    require('actions\\matchPair');
    require('actions\\editPoints');
    require('actions\\selectItem');
    require('actions\\selectLine');
    require('actions\\lineBreak');
    require('actions\\mergeLines');
    require('actions\\toggleComment');
    require('actions\\splitJoinTag');
    require('actions\\removeTag');
    require('actions\\evaluateMath');
    require('actions\\increment_decrement');
    require('actions\\base64');
    require('actions\\reflectCSSValue');
    require('actions\\updateImageSize');
    require('bootstrap').loadSystemSnippets(require('file').read(context.Root + '\\Assets\\snippets.json'));
    context.startup(JSON.stringify(require('actions').getMenu()));
});
