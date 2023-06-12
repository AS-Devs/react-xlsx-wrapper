"use strict";

var __importDefault = void 0 && (void 0).__importDefault || function (mod) {
  return mod && mod.__esModule ? mod : {
    "default": mod
  };
};
Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.ExcelColumn = exports.ExcelSheet = void 0;
var ExcelFileNew_1 = __importDefault(require("./ExcelPlugin/components/ExcelFileNew"));
var ExcelSheet_1 = __importDefault(require("./ExcelPlugin/elements/ExcelSheet"));
exports.ExcelSheet = ExcelSheet_1.default;
var ExcelColumn_1 = __importDefault(require("./ExcelPlugin/elements/ExcelColumn"));
exports.ExcelColumn = ExcelColumn_1.default;
ExcelFileNew_1.default.ExcelSheet = ExcelSheet_1.default;
ExcelFileNew_1.default.ExcelColumn = ExcelColumn_1.default;
exports.default = ExcelFileNew_1.default;