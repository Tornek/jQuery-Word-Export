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
				'mso-page-orientation': 'portrait',
				'mso-header': 'h1',
				'mso-footer': 'f1'
			},
			'div.Section1': {page: 'Section1'},
			'table#systemElems': {
				margin: '0in 0in 0in 9in'
			}
		},
		imageOptions: {
			maxWidth: 624
		}
	};
	
	$.fn.wordExport = function(params){
		params = $.extend(true, defaultParam, params);
		
		var mhtml = {
				top: "Mime-Version: 1.0\nContent-Type: Multipart/related; boundary=\"----=_NextPart_ZROIIZO.ZCZYUACXV.ZARTUI\"\n\n------=_NextPart_ZROIIZO.ZCZYUACXV.ZARTUI\nContent-Type: text/html; charset=\"utf-8\"\nContent-Location: " + location.href + "\n\n<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>\n_html_</html>",
				head: "<head>\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\n<style>\n<!--\n_styles_-->\n</style>\n</head>\n",
				body: "<body><div class='Section1'>_body_</div><table id='systemElems' border='1' cellspacing='0' cellpadding='0'><tr><td>_h1_</td></tr><tr><td>_f1_</td></tr></table></body>"
			},
			// Clone selected element before manipulating it
			$markup = $(this).clone(),
			stylesInLine = '';
		
		// Remove hidden elements from the output
		$markup.each(function(){
			var self = $(this);
			
			if(self.is(':hidden'))
				self.remove();
		});
		
		if($markup.find('#h1').length){
			var $header = $markup.find('#h1');
			
			mhtml.body = mhtml.body.replace("_h1_", $header.get(0).outerHTML);
			$header.remove();
		}
		if($markup.find('#f1').length){
			var $footer = $markup.find('#f1');
			
			mhtml.body = mhtml.body.replace("_f1_", $footer.get(0).outerHTML);
			$footer.remove();
		}
		
		// Embed all images using Data URLs
		var images = [],
			img = $markup.find('img'),
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
			
			mhtmlBottom += "------=_NextPart_ZROIIZO.ZCZYUACXV.ZARTUI\n";
			mhtmlBottom += "Content-Location: " + images[i].location + "\n";
			mhtmlBottom += "Content-Type: " + images[i].type + "\n";
			mhtmlBottom += "Content-Transfer-Encoding: " + images[i].encoding + "\n\n";
			mhtmlBottom += images[i].data + "\n\n";
		}
		//The bottom of bound
		mhtmlBottom += "------=_NextPart_ZROIIZO.ZCZYUACXV.ZARTUI--";
		
		$.each(params.styles, function(selector){
			stylesInLine += selector + ' {\n';
			
			$.each(params.styles[selector], function(item){
				stylesInLine += item + ': ' + params.styles[selector][item] + ';\n';
			});
			stylesInLine += '}\n';
		});
		
		// Aggregate parts of the file together
		var fileContent = mhtml.top.replace("_html_", mhtml.head.replace("_styles_", stylesInLine) + mhtml.body.replace("_body_", $markup.html())) + mhtmlBottom,
			// Create a Blob with the file contents
			blob = new Blob([fileContent], {
				type: "application/msword;charset=utf-8"
			});
		
		saveAs(blob, params.fileName + ".doc");
	};
})(jQuery);