/**
 * Install
 */
function onInstall() {
  onOpen();
}

/**
 * Create menu
 */
function onOpen() {
  var ui = SlidesApp
    .getUi()
    .createAddonMenu()
    .addItem('Launch ShapeApp', 'opensidebar')
    .addItem('About', 'about')
    .addToUi();
}

/**
 * Open sidebar
 */
function opensidebar() {
  var sidebarUi = HtmlService.createTemplateFromFile('sidebar').evaluate().setTitle('ShapeApp');
  SlidesApp.getUi().showSidebar(sidebarUi);
}

/**
 * Open about dialog
 */
function about() {
  var html = HtmlService.createHtmlOutputFromFile('about').setWidth(640).setHeight(480).setTitle('About');
  SlidesApp.getUi().showModalDialog(html, "About");
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function showDialog(e) {
  SlidesApp.getUi().alert(e);
}

/**
 * Create shape.<br>
 * @param {Object} Object Object
 * @return {Object} Return Object
 */
function createShape(p) {
    return new ShapeApp().createShape(p);
}

/**
 * Get selected shape.<br>
 * @return {Object} Return Object
 */
function getSelectedShape() {
    return new ShapeApp().getSelectedShape();
}

/**
 * Update shape.<br>
 * @param {Object} Object Object
 * @return {Object} Return Object
 */
function updateShape(p) {
    return new ShapeApp().updateShape(p);
}

/**
 * Arrange shapes.<br>
 * @param {Object} Object Object
 * @return {Object} Return Object
 */
function arrangeShapes(p) {
    return new ShapeApp().arrangeShapes(p);
}
;
(function(r) {
  var ShapeApp;
  ShapeApp = (function() {
    var arrangeShapesDo, axsisSort, copyShape, doArrange, getShapeByObjectId;

    ShapeApp.name = "ShapeApp";

    function ShapeApp() {
      this.slide = SlidesApp.getActivePresentation();
      this.presentationId = this.slide.getId();
      this.ui = SlidesApp.getUi();
    }

    ShapeApp.prototype.createShape = function(p_) {
      var e, f, j, len, objId, pageObjectId, slides;
      try {
        pageObjectId = this.slide.getSelection().getCurrentPage().getObjectId();
        slides = this.slide.getSlides();
        objId = "";
        for (j = 0, len = slides.length; j < len; j++) {
          e = slides[j];
          if (e.getObjectId() === pageObjectId) {
            f = e.insertShape(SlidesApp.ShapeType[p_.select_type], p_.translateX && p_.translateX > 0 ? p_.translateX : 0, p_.translateY && p_.translateY > 0 ? p_.translateY : 0, p_.width && p_.width > 0 ? p_.width : 100, p_.height && p_.height > 0 ? p_.height : 100);
            f.setRotation(p_.rotation ? p_.rotation : 0);
            copyShape.call(this, f, p_.number);
            break;
          }
        }
      } catch (error) {
        e = error;
        this.ui.alert(e);
        return "Error.";
      }
    };

    ShapeApp.prototype.getSelectedShape = function() {
      var e, pageObjectId, params, select, selectType, selected;
      select = this.slide.getSelection();
      pageObjectId = select.getCurrentPage().getObjectId();
      selectType = select.getSelectionType();
      if (selectType.toString() === "PAGE_ELEMENT") {
        selected = select.getPageElementRange().getPageElements();
        if (selected.length === 1) {
          try {
            params = {
              presentationId: this.presentationId,
              pageObjectId: pageObjectId,
              objectId: selected[0].getObjectId()
            };
            r = getShapeByObjectId.call(this, params);
            return {
              shapeType: r.shape ? r.shape.shapeType : "Not shape type",
              width: selected[0].getWidth(),
              height: selected[0].getHeight(),
              rotation: selected[0].getRotation(),
              translateX: selected[0].getLeft(),
              translateY: selected[0].getTop(),
              objectId: r.objectId
            };
          } catch (error) {
            e = error;
            this.ui.alert(e);
            return "Error.";
          }
        } else if (selected.length > 1) {
          this.ui.alert("Please select one shape.");
          return "Error.";
        }
      } else {
        this.ui.alert("Please select one shape.");
        return "Error.";
      }
    };

    ShapeApp.prototype.updateShape = function(p_) {
      var e, f, j, l, len, len1, pe, slides;
      if (p_.width <= 0 || p_.height <= 0) {
        this.ui.alert("Please set Width and Height more than zero.");
        return "Error.";
      }
      try {
        slides = this.slide.getSlides();
        for (j = 0, len = slides.length; j < len; j++) {
          e = slides[j];
          pe = e.getPageElements();
          for (l = 0, len1 = pe.length; l < len1; l++) {
            f = pe[l];
            if (f.getObjectId() === p_.objectId) {
              f.setWidth(p_.width);
              f.setHeight(p_.height);
              f.setLeft(p_.translateX);
              f.setTop(p_.translateY);
              f.setRotation(p_.rotation);
              copyShape.call(this, f, p_.number);
              return null;
            }
          }
        }
        this.ui.alert("Not found objectId: " + p_.objectId);
        return "Error.";
      } catch (error) {
        e = error;
        this.ui.alert(e);
        return "Error.";
      }
    };

    ShapeApp.prototype.arrangeShapes = function(p_) {
      var data, e, j, len, objs, pageLen, select, selected, temp, totalHeight, totalObjLen, totalWidth;
      select = this.slide.getSelection().getPageElementRange();
      if (!select) {
        this.ui.alert("Select shapes.");
        return "Error.";
      }
      selected = select.getPageElements();
      pageLen = p_.axsisX ? this.slide.getPageWidth() : this.slide.getPageHeight();
      objs = [];
      totalWidth = 0;
      totalHeight = 0;
      for (j = 0, len = selected.length; j < len; j++) {
        e = selected[j];
        temp = {
          object: e,
          objectId: e.getObjectId(),
          width: e.getWidth(),
          height: e.getHeight(),
          translateX: e.getLeft(),
          translateY: e.getTop()
        };
        objs.push(temp);
        totalWidth += temp.width;
        totalHeight += temp.height;
      }
      data = axsisSort.call(this, p_.axsisX, objs);
      totalObjLen = p_.axsisX ? totalWidth : totalHeight;
      arrangeShapesDo.call(this, p_, data, totalObjLen, pageLen);
    };

    arrangeShapesDo = function(p_, data, totalObjLen, pageLen) {
      var sep;
      sep = 0;
      if (data.length === 1 && !p_.edgeOpen) {
        sep = 0;
      } else {
        sep = p_.edgeOpen ? (pageLen - totalObjLen) / (data.length + 1) : (pageLen - totalObjLen) / (data.length - 1);
      }
      if (sep < 0) {
        this.ui.alert("Total length of selected objects is larger than the page width.");
        return "Error.";
      }
      doArrange.call(this, p_.edgeOpen ? 1 : 0, p_.axsisX ? "width" : "height", data, sep);
    };

    doArrange = function(f, axsis, data, sep) {
      var e, i, j, len, objLen;
      objLen = 0;
      for (i = j = 0, len = data.length; j < len; i = ++j) {
        e = data[i];
        if (axsis === "width") {
          e.object.setLeft((i + f) * sep + objLen);
        } else {
          e.object.setTop((i + f) * sep + objLen);
        }
        objLen += e[axsis];
      }
    };

    axsisSort = function(axsis, objs) {
      if (axsis) {
        objs.sort(function(x, y) {
          if (x.translateX < y.translateX) {
            return -1;
          }
          if (x.translateX > y.translateX) {
            return 1;
          }
          return 0;
        });
      } else {
        objs.sort(function(x, y) {
          if (x.translateY < y.translateY) {
            return -1;
          }
          if (x.translateY > y.translateY) {
            return 1;
          }
          return 0;
        });
      }
      return objs;
    };

    getShapeByObjectId = function(p_) {
      var e, j, len;
      r = Slides.Presentations.Pages.get(p_.presentationId, p_.pageObjectId).pageElements;
      for (j = 0, len = r.length; j < len; j++) {
        e = r[j];
        if (e.objectId === p_.objectId) {
          return e;
        }
      }
      return null;
    };

    copyShape = function(f, number) {
      var cpe, j, k, posX, posY, ref, results;
      if (number && number > 1) {
        posX = f.getLeft();
        posY = f.getTop();
        results = [];
        for (k = j = 0, ref = number - 2; 0 <= ref ? j <= ref : j >= ref; k = 0 <= ref ? ++j : --j) {
          cpe = f.duplicate();
          posX += 10;
          posY += 10;
          cpe.setLeft(posX);
          results.push(cpe.setTop(posY));
        }
        return results;
      }
    };

    return ShapeApp;

  })();
  return r.ShapeApp = ShapeApp;
})(this);
