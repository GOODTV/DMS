/**
 * @license Copyright (c) 2003-2013, CKSource - Frederico Knabben. All rights reserved.
 * For licensing, see LICENSE.html or http://ckeditor.com/license
 */

CKEDITOR.editorConfig = function( config ) {
	// Define changes to default configuration here. For example:
    //config.language = 'zh-cn';
    // config.uiColor = '#AADC6E';
    config.skin = 'moonocolor_v1.1';

    config.toolbar = 'MyToolbar';
    config.toolbar_MyToolbar =
    [
        ['Source', 'NewPage', 'Preview'],
        ['Bold', 'Italic', 'Underline', 'Strike'],
        ['NumberedList', 'BulletedList', '-', 'Outdent', 'Indent'],
        ['JustifyLeft', 'JustifyCenter', 'JustifyRight', 'JustifyBlock'],
        ['Maximize', 'SelectAll'],
        '/',
        ['Format', 'Font', 'FontSize'],
        ['TextColor', 'BGColor', 'Link', 'Unlink', 'Anchor'],
    //['Cut','Copy','Paste','PasteText','PasteFromWord'],
    //['Undo', 'Redo', '-', 'Find', 'Replace', '-', 'SelectAll', 'RemoveFormat','Flash',,'Smiley','SpecialChar'],               
        ['Image', 'Flash', 'Table', 'HorizontalRule', 'RemoveFormat']
    ];
	config.font_names = "新細明體;標楷體;微軟正黑體;" +config.font_names ;	
};
