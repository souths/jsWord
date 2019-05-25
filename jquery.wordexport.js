/**
 * 此源码有两处修改,1、源码未修改前导出的word是以web视图的方式打开，修改后Word以“页面方式”打开
 * 修改打开视图的方式参考https://blog.csdn.net/fengshuiyue/article/details/73321190
 * 2、修改了图表转换成canvas导出图片
 */
if (typeof jQuery !== "undefined" && typeof saveAs !== "undefined") {
  (function($) {
    $.fn.wordExport = function(fileName) {
      fileName = typeof fileName !== 'undefined' ? fileName : "jQuery-Word-Export";
      var page = {
        mhtml: {
          top: "Mime-Version: 1.0\nContent-Base: " + location.href + "\nContent-Type: Multipart/related; boundary=\"NEXT.ITEM-BOUNDARY\";type=\"text/html\"\n\n--NEXT.ITEM-BOUNDARY\nContent-Type: text/html; charset=\"utf-8\"\nContent-Location: " + location.href + "\n\n<!DOCTYPE html>\n<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\">"+
          "\n_html_</html>",
          head: "<head>\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">"+
          "\n<!--[if gte mso 9]>"+
          "<xml>"+
          "\n<w:WordDocument>"+
          "\n<w:View>Print</w:View>"+
          "\n<w:GrammarState>Clean</w:GrammarState>"+
          "\n<w:TrackMoves>false</w:TrackMoves>"+
          "\n<w:TrackFormatting/>"+
          "\n<w:ValidateAgainstSchemas/>"+
          "\n<w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>"+
          "\n<w:IgnoreMixedContent>false</w:IgnoreMixedContent>"+
          "\n<w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>"+
          "\n<w:DoNotPromoteQF/>"+
          "\n<w:LidThemeOther>EN-US</w:LidThemeOther>"+
          "\n<w:LidThemeAsian>ZH-CN</w:LidThemeAsian>"+
          "\n<w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript>"+
          "\n<w:Compatibility>"+
          "\n<w:BreakWrappedTables/>"+
          "\n<w:SnapToGridInCell/>"+
          "\n<w:WrapTextWithPunct/>"+
          "\n<w:UseAsianBreakRules/>"+
          "\n<w:DontGrowAutofit/>"+
          "\n<w:SplitPgBreakAndParaMark/>"+
          "\n<w:DontVertAlignCellWithSp/>"+
          "\n<w:DontBreakConstrainedForcedTables/>"+
          "\n<w:DontVertAlignInTxbx/>"+
          "\n<w:Word11KerningPairs/>"+
          "\n<w:CachedColBalance/>"+
          "\n<w:UseFELayout/>"+
          "\n</w:Compatibility>"+
          "\n<w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>"+
          "\n<m:mathPr>"+
          "\n<m:mathFont m:val=\"Cambria Math\"/>"+
          "\n<m:brkBin m:val=\"before\"/>"+
          "\n<m:brkBinSub m:val=\"--\"/>"+
          "\n<m:smallFrac m:val=\"off\"/>"+
          "\n<m:dispDef/>"+
          "\n<m:lMargin m:val=\"0\"/>"+
          "\n<m:rMargin m:val=\"0\"/>"+
          "\n<m:defJc m:val=\"centerGroup\"/>"+
          "\n<m:wrapIndent m:val=\"1440\"/>"+
          "\n<m:intLim m:val=\"subSup\"/>"+
          "\n<m:naryLim m:val=\"undOvr\"/>"+
          "\n</m:mathPr></w:WordDocument>"+
          "\n</xml>"+
          "\n<![endif]-->"+
          "\n<style>\n_styles_\n</style>\n</head>\n",
          body: "<body>_body_</body>"
        }
      };
      var options = {
        maxWidth: 624
      };
      // Clone selected element before manipulating it
      var markup = $(this).clone();

      // Remove hidden elements from the output
      markup.each(function() {
        var self = $(this);
        if (self.is(':hidden'))
          self.remove();
      });

      // Embed all images using Data URLs
      var images = Array();
       var img = markup.find('img');
     // var img = new Image();
      for (var i = 0; i < img.length; i++) {
        // Calculate dimensions of output image
       // var w = Math.min(img[i].width, options.maxWidth);
      //  var h = img[i].height * (w / img[i].width);
        // Create canvas for converting image to data URL
      //  var canvas = document.createElement("CANVAS");
     //   canvas.width = w;
      //  canvas.height = h;
        // Draw image to canvas
       // var context = canvas.getContext('2d');
      //  context.drawImage(img[i], 0, 0, w, h);
        // Get data URL encoding of image
       // var uri = canvas.toDataURL("image/png");
        var uri = img[i].src
        /**
         * var uri = img[i].src 这个代码不属于jquery.wordexport.js
         * html页面的echart图表都已经转换成img的base64，
         * 此处的canvas是为了把页面的图片转成img的base64,所以注释此处的canvas,
         * 但是下面需要用到图片的src的base64，所以取出图片的base64
         *
         */
        $(img[i]).attr("src", img[i].src);
      //  img[i].width = w;
      // img[i].height = h;
        // Save encoded image to array
        images[i] = {
          type: uri.substring(uri.indexOf(":") + 1, uri.indexOf(";")),
          encoding: uri.substring(uri.indexOf(";") + 1, uri.indexOf(",")),
          location: $(img[i]).attr("src"),
          data: uri.substring(uri.indexOf(",") + 1)
        };
      }

      // Prepare bottom of mhtml file with image data
      var mhtmlBottom = "\n";
      for (var i = 0; i < images.length; i++) {
        mhtmlBottom += "--NEXT.ITEM-BOUNDARY\n";
        mhtmlBottom += "Content-Location: " + images[i].location + "\n";
        mhtmlBottom += "Content-Type: " + images[i].type + "\n";
        mhtmlBottom += "Content-Transfer-Encoding: " + images[i].encoding + "\n\n";
        mhtmlBottom += images[i].data + "\n\n";
      }
      mhtmlBottom += "--NEXT.ITEM-BOUNDARY--";

      //TODO: load css from included stylesheet
      var styles = "@page WordSection1 {size: 595.3pt 841.9pt;margin: 1.5cm 2.0cm 1.5cm 2.0cm;mso-header-margin: 42.55pt;mso-footer-margin: 49.6pt;mso-paper-source: 0;}div.WordSection1 {page: WordSection1}";

      // Aggregate parts of the file together
      var fileContent = page.mhtml.top.replace("_html_", page.mhtml.head.replace("_styles_", styles) + page.mhtml.body.replace("_body_", markup.html())) + mhtmlBottom;

      // Create a Blob with the file contents
      var blob = new Blob([fileContent], {
        type: "application/msword;charset=utf-8"
      });
      saveAs(blob, fileName + ".doc");
    };
  })(jQuery);
} else {
  if (typeof jQuery === "undefined") {
    console.error("jQuery Word Export: missing dependency (jQuery)");
  }
  if (typeof saveAs === "undefined") {
    console.error("jQuery Word Export: missing dependency (FileSaver.js)");
  }
}
