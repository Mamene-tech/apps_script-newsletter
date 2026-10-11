/**
 * EmailTemplate.js / EmailTemplate.gs
 * Generates 100% email-client compatible responsive HTML email (no Flexbox).
 * Built with standard nested <table> structures, align="left", and CSS media queries.
 * 
 * Features:
 * - 100% Optional image captions (renders NOTHING if empty)
 * - Highly distinct section cards with crisp borders and subtle shadows
 */

var EmailTemplate = {
  buildHtml: function(data, recipientEmail, recipientName, webAppUrl) {
    var primaryColor = "#298DC0";
    var recipient = recipientEmail || "vas-email@domena.cz";
    var name = recipientName ? recipientName.trim() : "";

    // Helper to replace personalization placeholders
    function replacePlaceholders(text) {
      if (!text) return "";
      return text
        .replace(/@jméno|@jmeno|@name|\{jméno\}|\{jmeno\}|\{name\}|\[jméno\]|\[jmeno\]|\[name\]/gi, name || "")
        .replace(/@email|\{email\}|\[email\]/gi, recipient || "");
    }

    // Dynamic Greeting Logic: strictly respects user's custom greeting text
    var rawGreeting = (data.greeting || "Dobrý den @jméno,").trim();
    var greeting = "";
    var hasPlaceholder = /@jméno|@jmeno|@name|\{jméno\}|\{jmeno\}|\{name\}|\[jméno\]|\[jmeno\]|\[name\]/i.test(rawGreeting);

    if (hasPlaceholder) {
      if (name) {
        greeting = replacePlaceholders(rawGreeting);
      } else {
        // Recipient has no name in sheet: clean up placeholder cleanly
        greeting = rawGreeting
          .replace(/@jméno|@jmeno|@name|\{jméno\}|\{jmeno\}|\{name\}|\[jméno\]|\[jmeno\]|\[name\]/gi, "")
          .replace(/,\s*,/g, ",")
          .replace(/\s+,/g, ",")
          .replace(/\s+/g, " ")
          .trim();
      }
    } else {
      // User did not write @jmeno explicitly
      if (name) {
        if (rawGreeting.indexOf(",") !== -1) {
          var commaIdx = rawGreeting.indexOf(",");
          var before = rawGreeting.substring(0, commaIdx).trim();
          var after = rawGreeting.substring(commaIdx + 1).trim();
          greeting = before + " " + name + (after ? (", " + after) : ",");
        } else {
          greeting = rawGreeting + " " + name;
        }
      } else {
        greeting = rawGreeting;
      }
    }

    var unsubUrl = webAppUrl ? 
      (webAppUrl + (webAppUrl.indexOf("?") === -1 ? "?" : "&") + "action=unsub&email=" + encodeURIComponent(recipient)) : "#";

    // Helper: Render individual content element
    function renderElement(el) {
      if (!el) return "";
      
      if (el.type === "heading") {
        return '<h3 style="color: ' + primaryColor + '; margin: 0 0 10px 0; font-size: 19px; font-weight: 700; font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif;">' + (el.text || "") + '</h3>';
      }
      
      if (el.type === "text") {
        var align = el.align || "left";
        var textHtml = (el.text || el.html || "").replace(/\n/g, '<br>');
        return '<div style="color: #4a5568; font-size: 15px; line-height: 1.65; margin-bottom: 12px; text-align: ' + align + '; font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif;">' + textHtml + '</div>';
      }
      
      if (el.type === "image") {
        var imgUrl = el.url || el.imageUrl || "";
        if (!imgUrl) return "";
        
        // STRICTLY OPTIONAL CAPTION: Only render if not empty!
        var rawCap = (el.caption || el.imageCaption || "").trim();
        var captionTag = rawCap !== "" ? 
          ('<div style="font-size: 12px; color: #718096; margin-top: 6px; font-style: italic; text-align: center;">' + rawCap + '</div>') : "";

        return [
          '<div style="text-align: center; margin: 10px 0;">',
          '<img src="' + imgUrl + '" class="mobile-full-img" style="max-width: 100%; height: auto; border-radius: 6px; border: 1px solid #e2e8f0; display: inline-block;">',
          captionTag,
          '</div>'
        ].join('');
      }

      if (el.type === "button") {
        var btnAlign = el.align || "center";
        return [
          '<div style="text-align: ' + btnAlign + '; margin: 22px 0;">',
          '<a href="' + (el.url || "#") + '" target="_blank" style="background-color: ' + primaryColor + '; color: #ffffff; text-decoration: none; padding: 13px 26px; font-size: 15px; font-weight: 700; border-radius: 6px; display: inline-block; box-shadow: 0 2px 4px rgba(41,141,192,0.25); font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif;">',
          (el.text || "Klikněte zde"),
          '</a>',
          '</div>'
        ].join('');
      }

      if (el.type === "highlight") {
        return [
          '<div style="background-color: #f4f8fb; border-left: 4px solid ' + primaryColor + '; padding: 15px 20px; margin: 16px 0; border-radius: 4px; font-size: 15px; color: #2d3748; line-height: 1.55; font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif;">',
          (el.text || el.html || "").replace(/\n/g, '<br>'),
          '</div>'
        ].join('');
      }

      if (el.type === "divider") {
        return '<div style="border-top: 1px solid #e2e8f0; margin: 25px 0;"></div>';
      }

      return "";
    }

    // Build Distinct Section Cards using pure nested <table>
    var sectionsHtml = "";
    if (data.sections && data.sections.length > 0) {
      for (var i = 0; i < data.sections.length; i++) {
        var sec = data.sections[i];
        if (!sec) continue;

        var secBorder = sec.hasBorder !== false;
        // Visually distinct section styling
        var borderStyle = secBorder ? 
          'background-color: #ffffff; border-left: 4px solid ' + primaryColor + '; border-top: 1px solid #d2dbe5; border-right: 1px solid #d2dbe5; border-bottom: 2px solid #cbd5e1; border-radius: 8px;' : 
          '';

        if (sec.cols && sec.cols.length > 0) {
          var numCols = sec.cols.length;
          
          if (numCols === 1) {
            // 1 column full width
            var colContent = "";
            for (var c = 0; c < sec.cols[0].length; c++) {
              colContent += renderElement(sec.cols[0][c]);
            }
            sectionsHtml += [
              '<table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin: 24px 0; ' + borderStyle + '">',
              '<tr><td style="' + (secBorder ? 'padding: 24px;' : 'padding: 5px;') + '">',
              colContent,
              '</td></tr></table>'
            ].join('');

          } else if (numCols === 2) {
            // 2 Columns
            var leftContent = "";
            for (var l = 0; l < sec.cols[0].length; l++) {
              leftContent += renderElement(sec.cols[0][l]);
            }
            var rightContent = "";
            for (var r = 0; r < sec.cols[1].length; r++) {
              rightContent += renderElement(sec.cols[1][r]);
            }

            sectionsHtml += [
              '<table border="0" cellpadding="0" cellspacing="0" width="100%" class="responsive-card" style="margin: 24px 0; ' + borderStyle + '">',
              '<tr><td style="' + (secBorder ? 'padding: 24px;' : 'padding: 5px;') + '">',
              '  <table border="0" cellpadding="0" cellspacing="0" width="100%">',
              '  <tr>',
              '    <td class="stack-column" width="48%" valign="top" style="padding-bottom: 10px;">' + leftContent + '</td>',
              '    <td class="stack-spacer" width="4%">&nbsp;</td>',
              '    <td class="stack-column" width="48%" valign="top" style="padding-bottom: 10px;">' + rightContent + '</td>',
              '  </tr>',
              '  </table>',
              '</td></tr></table>'
            ].join('');

          } else if (numCols === 3) {
            // 3 Columns
            var col1Html = "";
            for (var a = 0; a < sec.cols[0].length; a++) col1Html += renderElement(sec.cols[0][a]);
            var col2Html = "";
            for (var b = 0; b < sec.cols[1].length; b++) col2Html += renderElement(sec.cols[1][b]);
            var col3Html = "";
            for (var k = 0; k < sec.cols[2].length; k++) col3Html += renderElement(sec.cols[2][k]);

            sectionsHtml += [
              '<table border="0" cellpadding="0" cellspacing="0" width="100%" class="responsive-card" style="margin: 24px 0; ' + borderStyle + '">',
              '<tr><td style="' + (secBorder ? 'padding: 24px;' : 'padding: 5px;') + '">',
              '  <table border="0" cellpadding="0" cellspacing="0" width="100%">',
              '  <tr>',
              '    <td class="stack-column-3" width="31%" valign="top" style="padding-bottom: 10px;">' + col1Html + '</td>',
              '    <td class="stack-spacer" width="3.5%">&nbsp;</td>',
              '    <td class="stack-column-3" width="31%" valign="top" style="padding-bottom: 10px;">' + col2Html + '</td>',
              '    <td class="stack-spacer" width="3.5%">&nbsp;</td>',
              '    <td class="stack-column-3" width="31%" valign="top" style="padding-bottom: 10px;">' + col3Html + '</td>',
              '  </tr>',
              '  </table>',
              '</td></tr></table>'
            ].join('');
          }
        }
      }
    }

    return [
      '<!DOCTYPE html>',
      '<html>',
      '<head>',
      '<meta charset="utf-8">',
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
      '<title>' + (data.subject || "Newsletter") + '</title>',
      '<style type="text/css">',
      '  body, table, td, a { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }',
      '  table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; }',
      '  img { -ms-interpolation-mode: bicubic; border: 0; outline: none; text-decoration: none; }',
      '  ',
      '  @media screen and (max-width: 640px) {',
      '    .email-container { width: 100% !important; max-width: 100% !important; }',
      '    .stack-column { display: block !important; width: 100% !important; max-width: 100% !important; box-sizing: border-box !important; padding: 0 0 15px 0 !important; }',
      '    .stack-column-3 { display: block !important; width: 100% !important; max-width: 100% !important; box-sizing: border-box !important; padding: 0 0 15px 0 !important; }',
      '    .stack-spacer { display: none !important; width: 0 !important; height: 0 !important; }',
      '    .mobile-full-img { width: 100% !important; max-width: 100% !important; height: auto !important; }',
      '    .mobile-padding { padding: 20px 16px !important; }',
      '    .mobile-header-padding { padding: 22px 18px !important; }',
      '  }',
      '</style>',
      '</head>',
      '<body style="margin: 0; padding: 20px 8px; background-color: #edf2f7; font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif; -webkit-font-smoothing: antialiased;">',
      '<center>',
      '<table border="0" cellpadding="0" cellspacing="0" width="100%" class="email-container" style="max-width: 840px; text-align: left; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.06); border: 1px solid #e2e8f0;">',
      
      // Header Banner
      '<tr>',
      '<td class="mobile-header-padding" style="background-color: ' + primaryColor + '; padding: 28px 36px; text-align: left;">',
      '<h1 style="color: #ffffff; margin: 0; font-size: 25px; font-weight: 700; letter-spacing: -0.3px; font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif;">' + (data.headerTitle || "SHM Uhříněves - Kolovraty newsletter") + '</h1>',
      '</td>',
      '</tr>',

      // Body Card
      '<tr>',
      '<td class="mobile-padding" style="padding: 34px;">',
      
      '<p style="color: #2d3748; font-size: 16px; line-height: 1.6; margin: 0 0 14px 0; font-weight: 600;">' + greeting + '</p>',
      data.introText ? ('<div style="color: #4a5568; font-size: 15px; line-height: 1.65; margin-bottom: 22px;">' + replacePlaceholders(data.introText).replace(/\n/g, '<br>') + '</div>') : '',
      
      data.highlightBoxText ? [
        '<div style="background-color: #f4f8fb; border-left: 4px solid ' + primaryColor + '; padding: 15px 20px; margin: 20px 0; border-radius: 4px; font-size: 15px; color: #2d3748; line-height: 1.55;">',
        data.highlightBoxText.replace(/\n/g, '<br>'),
        '</div>'
      ].join('') : '',

      '<div style="border-top: 1px solid #e2e8f0; margin: 26px 0;"></div>',

      sectionsHtml,

      data.signoff ? ('<div style="color: #4a5568; font-size: 15px; line-height: 1.65; margin-top: 28px;">' + data.signoff.replace(/\n/g, '<br>') + '</div>') : '',
      
      '<div style="border-top: 1px solid #e2e8f0; margin: 34px 0 16px 0;"></div>',
      '<div style="text-align: center; color: #a0aec0; font-size: 12px; line-height: 1.5;">',
      'Tento email byl odeslán na adresu <strong>' + recipient + '</strong>.<br>',
      'Pokud si již nepřejete dostávat další emaily, můžete se <a href="' + unsubUrl + '" style="color: ' + primaryColor + '; text-decoration: underline; font-weight: 600;">odhlásit z odběru</a>.',
      '</div>',

      '</td>',
      '</tr>',
      '</table>',
      '</center>',
      '</body>',
      '</html>'
    ].join('\n');
  }
};