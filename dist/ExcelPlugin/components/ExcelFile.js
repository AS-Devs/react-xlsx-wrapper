"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");
var _classCallCheck2 = _interopRequireDefault(require("@babel/runtime/helpers/classCallCheck"));
var _createClass2 = _interopRequireDefault(require("@babel/runtime/helpers/createClass"));
var _assertThisInitialized2 = _interopRequireDefault(require("@babel/runtime/helpers/assertThisInitialized"));
var _inherits2 = _interopRequireDefault(require("@babel/runtime/helpers/inherits"));
var _possibleConstructorReturn2 = _interopRequireDefault(require("@babel/runtime/helpers/possibleConstructorReturn"));
var _getPrototypeOf2 = _interopRequireDefault(require("@babel/runtime/helpers/getPrototypeOf"));
var _defineProperty2 = _interopRequireDefault(require("@babel/runtime/helpers/defineProperty"));
function _createSuper(Derived) { var hasNativeReflectConstruct = _isNativeReflectConstruct(); return function _createSuperInternal() { var Super = (0, _getPrototypeOf2.default)(Derived), result; if (hasNativeReflectConstruct) { var NewTarget = (0, _getPrototypeOf2.default)(this).constructor; result = Reflect.construct(Super, arguments, NewTarget); } else { result = Super.apply(this, arguments); } return (0, _possibleConstructorReturn2.default)(this, result); }; }
function _isNativeReflectConstruct() { if (typeof Reflect === "undefined" || !Reflect.construct) return false; if (Reflect.construct.sham) return false; if (typeof Proxy === "function") return true; try { Boolean.prototype.valueOf.call(Reflect.construct(Boolean, [], function () {})); return true; } catch (e) { return false; } }
var __importDefault = void 0 && (void 0).__importDefault || function (mod) {
  return mod && mod.__esModule ? mod : {
    "default": mod
  };
};
Object.defineProperty(exports, "__esModule", {
  value: true
});
var react_1 = __importDefault(require("react"));
var xlsx_js_style_1 = require("xlsx-js-style");
var DataUtil_1 = require("../utils/DataUtil");
var ExcelFile = /*#__PURE__*/function (_react_1$default$Comp) {
  (0, _inherits2.default)(ExcelFile, _react_1$default$Comp);
  var _super = _createSuper(ExcelFile);
  function ExcelFile() {
    var _this;
    (0, _classCallCheck2.default)(this, ExcelFile);
    for (var _len = arguments.length, args = new Array(_len), _key = 0; _key < _len; _key++) {
      args[_key] = arguments[_key];
    }
    _this = _super.call.apply(_super, [this].concat(args));
    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), "state", {
      fileName: "Download",
      fileExtension: "xlsx",
      hideElement: false
    });
    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), "createSheetData", function (sheet) {
      var columns = sheet.props.children;
      var sheetData = [react_1.default.Children.map(columns, function (column) {
        return column.props.label;
      })];
      var data = sheet.props.data;
      if (!data) throw new Error("No data provided");
      data.forEach(function (row) {
        var sheetRow = [];
        react_1.default.Children.forEach(columns, function (column) {
          var getValue = function getValue(row) {
            return typeof column.props.value === 'function' ? column.props.value(row) : row[column.props.value];
          };
          var itemValue = getValue(row);
          sheetRow.push(isNaN(Number(itemValue)) ? itemValue || "" : itemValue);
        });
        sheetData.push(sheetRow);
      });
      return sheetData;
    });
    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), "download", function () {
      var wb = xlsx_js_style_1.utils.book_new();
      var fileName = _this.getFileName();
      var fileExtension = _this.getFileExtension();
      react_1.default.Children.forEach(_this.props.children, function (sheet) {
        var ws = {};
        var wsName = sheet.props.name || fileName.split(".")[0] || "Sheet1";
        if (typeof sheet.props.dataSet === "undefined" || sheet.props.dataSet.length === 0) {
          ws = (0, DataUtil_1.excelSheetFromAoA)(_this.createSheetData(sheet));
        } else {
          ws = (0, DataUtil_1.excelSheetFromDataSet)(sheet.props.dataSet, sheet.props.bigHeading, sheet.props.autoFilterForAllColumn);
        }
        // add worksheet to workbook
        xlsx_js_style_1.utils.book_append_sheet(wb, ws, wsName);
      });
      (0, xlsx_js_style_1.writeFile)(wb, fileName, {
        bookType: fileExtension,
        bookSST: true,
        type: "binary",
        cellStyles: true
      });
    });
    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), "getFileNameWithExtension", function (filename, extension) {
      return "".concat(filename, ".").concat(extension);
    });
    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), "getFileName", function () {
      var _this$state$fileName;
      if (_this.state.fileName === null || typeof _this.state.fileName !== "string") {
        throw new Error("Invalid file name provided");
      }
      return _this.getFileNameWithExtension((_this$state$fileName = _this.state.fileName) === null || _this$state$fileName === void 0 ? void 0 : _this$state$fileName.split(".")[0], _this.getFileExtension());
    });
    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), "getFileExtension", function () {
      if (_this.props.fileExtension) {
        return _this.state.fileExtension;
      }
      // Fall back to extension embedded in the filename
      var slugs = _this.state.fileName.split(".");
      if (slugs.length > 1) {
        return slugs[slugs.length - 1];
      }
      return _this.state.fileExtension;
    });
    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), "handleDownload", function () {
      _this.download();
    });
    return _this;
  }
  (0, _createClass2.default)(ExcelFile, [{
    key: "componentDidMount",
    value: function componentDidMount() {
      if (this.props.filename) {
        this.setState({
          fileName: this.props.filename
        });
      }
      if (this.props.fileExtension) {
        this.setState({
          fileExtension: this.props.fileExtension
        });
      }
      if (this.props.hideElement) {
        this.setState({
          hideElement: this.props.hideElement
        });
      }
    }
  }, {
    key: "render",
    value: function render() {
      var element = this.props.element;
      if (this.state.hideElement === true) {
        return null;
      } else {
        return react_1.default.createElement("span", {
          onClick: this.handleDownload
        }, element);
      }
    }
  }]);
  return ExcelFile;
}(react_1.default.Component);
exports.default = ExcelFile;