/* 
 * jQuery.fn.wordExport
 * @version 0.1 (2012-12-16)
 * 
 * @desc Convert html code to word file (.doc)
 * 
 * @requires jQuery >= 1.11.1.
 * @requires FileSaver.js
 */
(function($){
	var defaultParam = {
		fileName: 'jQuery-Word-Export',
		styles: {
			'@page': {
				size: '21cm 29.7cmt',
				margin: '1cm 1cm 1cm 1cm',
				'mso-page-orientation': 'portrait'
			},
			'div.Section1': {page: 'Section1'}
		},
		imageOptions: {
			maxWidth: 624
		}
	};
	
	$.fn.wordExport = function(params){
		params = $.extend(true, defaultParam, params);
		
		var mhtml = {
				top: "Mime-Version: 1.0\nContent-Base: " + location.href + "\nContent-Type: Multipart/related; boundary=\"NEXT.ITEM-BOUNDARY\";type=\"text/html\"\n\n--NEXT.ITEM-BOUNDARY\nContent-Type: text/html; charset=\"utf-8\"\nContent-Location: " + location.href + "\n\n<!DOCTYPE html>\n<html>\n_html_</html>",
				head: "<head>\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\n<style>\n_styles_\n</style>\n</head>\n",
				body: "<body>_body_</body>"
			},
			// Clone selected element before manipulating it
			markup = $(this).clone(),
			stylesInLine = '';
		
		// Remove hidden elements from the output
		markup.each(function(){
			var self = $(this);
			
			if(self.is(':hidden'))
				self.remove();
		});
		
		// Embed all images using Data URLs
		var images = [],
			img = markup.find('img'),
			// Prepare bottom of mhtml file with image data
			mhtmlBottom = "\n";
		
		for(var i = 0; i < img.length; i++){
			// Calculate dimensions of output image
			var w = Math.min(img[i].width, params.imageOptions.maxWidth),
				h = img[i].height * (w / img[i].width),
				// Create canvas for converting image to data URL
				canvas = document.createElement("CANVAS");
			
			canvas.width = w;
			canvas.height = h;
			// Draw image to canvas
			var context = canvas.getContext('2d');
			
			context.drawImage(img[i], 0, 0, w, h);
			// Get data URL encoding of image
			var uri = canvas.toDataURL("image/png");
			
			$(img[i]).attr("src", img[i].src);
			img[i].width = w;
			img[i].height = h;
			// Save encoded image to array
			images[i] = {
				type: uri.substring(uri.indexOf(":") + 1, uri.indexOf(";")),
				encoding: uri.substring(uri.indexOf(";") + 1, uri.indexOf(",")),
				location: $(img[i]).attr("src"),
				data: uri.substring(uri.indexOf(",") + 1)
			};
			
			mhtmlBottom += "--NEXT.ITEM-BOUNDARY\n";
			mhtmlBottom += "Content-Location: " + images[i].location + "\n";
			mhtmlBottom += "Content-Type: " + images[i].type + "\n";
			mhtmlBottom += "Content-Transfer-Encoding: " + images[i].encoding + "\n\n";
			mhtmlBottom += images[i].data + "\n\n";
		}
		mhtmlBottom += "--NEXT.ITEM-BOUNDARY--";
		
		$.each(params.styles, function(selector){
			stylesInLine += selector + ' {\n';
			
			$.each(params.styles[selector], function(item){
				stylesInLine += item + ': ' + params.styles[selector][item] + ';\n';
			});
			stylesInLine += '}\n';
		});
		
		// Aggregate parts of the file together
		var fileContent = mhtml.top.replace("_html_", mhtml.head.replace("_styles_", stylesInLine) + mhtml.body.replace("_body_", markup.html())) + mhtmlBottom,
			// Create a Blob with the file contents
			blob = new Blob([fileContent], {
				type: "application/msword;charset=utf-8"
			});
		saveAs(blob, params.fileName + ".doc");
	};
})(jQuery);