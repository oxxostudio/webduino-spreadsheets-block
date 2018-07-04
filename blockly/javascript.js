Blockly.JavaScript['sheet_init'] = function (block) {
  let value_name = Blockly.JavaScript.valueToCode(block, 'name', Blockly.JavaScript.ORDER_ATOMIC);
  let value_sheeturl = Blockly.JavaScript.valueToCode(block, 'sheetUrl', Blockly.JavaScript.ORDER_ATOMIC);
  let value_sheetname = Blockly.JavaScript.valueToCode(block, 'sheetName', Blockly.JavaScript.ORDER_ATOMIC);
  let code = value_name + ' = {};\n' +
    value_name + '.sheetUrl = ' + value_sheeturl + ';\n' +
    value_name + '.sheetName = ' + value_sheetname + ';\n';
  return code;
};

Blockly.JavaScript['sheet_write_normal'] = function (block) {
  let value_data = Blockly.JavaScript.valueToCode(block, 'data', Blockly.JavaScript.ORDER_ATOMIC);
  let variable_name = Blockly.JavaScript.variableDB_.getName(block.getFieldValue('NAME'), Blockly.Variables.NAME_TYPE);
  let value_range = Blockly.JavaScript.valueToCode(block, 'range', Blockly.JavaScript.ORDER_ATOMIC);
  let code = variable_name + '.type = "auto";\n' +
    variable_name + '.data = ' + value_data + ';\n' +
    variable_name + '.range = ' + value_range + ';\n' +
    'sheetWriteData(' + variable_name + ');\n';
  return code;
};

Blockly.JavaScript['sheet_write_easy'] = function (block) {
  let variable_name = Blockly.JavaScript.variableDB_.getName(block.getFieldValue('var_'), Blockly.Variables.NAME_TYPE);
  let dropdown_type = block.getFieldValue('type');
  let code;
  if (block.itemCount_ > 1) {
    code = '[';
    for (let n = 0; n < block.itemCount_; n++) {
      let val = Blockly.JavaScript.valueToCode(block, 'data_' + n) || '""';
      let newVal;
      if (val.indexOf(',') != -1) {
        newVal = "'" + val.replace(/,/g, '').replace(/\'/g, '').replace(/ /g, '') + "'";
      } else {
        newVal = val;
      }
      if (n < (block.itemCount_ - 1)) {
        code = code + newVal + ',';
      } else {
        code = code + newVal;
      }
    }
    code = variable_name + '.type = "' + dropdown_type + '";\n' +
      variable_name + '.data = ' + code + '];\n' +
      variable_name + '.range = "auto";\n' +
      'sheetWriteData(' + variable_name + ');\n';
  } else {
    let val = Blockly.JavaScript.valueToCode(block, 'data_' + 0) || '""';
    let newVal;
    if (val.indexOf(',') != -1) {
      newVal = "'" + val.replace(/,/g, '').replace(/\'/g, '').replace(/ /g, '') + "'";
    } else {
      newVal = val;
    }
    code = variable_name + '.type = "' + dropdown_type + '";\n' +
      variable_name + '.data = ' + newVal + ';\n' +
      variable_name + '.range = "auto";\n' +
      'sheetWriteData(' + variable_name + ');\n';
  }
  return code;
};

Blockly.JavaScript['sheet_read'] = function (block) {
  let variable_name = Blockly.JavaScript.variableDB_.getName(block.getFieldValue('name'), Blockly.Variables.NAME_TYPE);
  let statements_do = Blockly.JavaScript.statementToCode(block, 'do');
  let code = 'sheetReadData(' + variable_name + ',function(data){\n' +
    statements_do +
    '});\n';
  return code;
};

Blockly.JavaScript['sheet_read_data'] = function (block) {
  let value_range = Blockly.JavaScript.valueToCode(block, 'range', Blockly.JavaScript.ORDER_ATOMIC);
  let code = 'data.cell[' + value_range + ']';
  return [code, Blockly.JavaScript.ORDER_NONE];
};

Blockly.JavaScript['sheet_read_data_all'] = function (block) {
  let code = 'data.data';
  return [code, Blockly.JavaScript.ORDER_ATOMIC];
};

Blockly.JavaScript['sheet_read_data_last'] = function (block) {
  let dropdown_type = block.getFieldValue('type');
  let code = 'data.' + dropdown_type;
  return [code, Blockly.JavaScript.ORDER_NONE];
};

Blockly.JavaScript['sheet_delete_row'] = function (block) {
  let variable_name = Blockly.JavaScript.variableDB_.getName(block.getFieldValue('name'), Blockly.Variables.NAME_TYPE);
  let value_num = Blockly.JavaScript.valueToCode(block, 'num', Blockly.JavaScript.ORDER_ATOMIC);
  let value_delete_num = Blockly.JavaScript.valueToCode(block, 'delete_num', Blockly.JavaScript.ORDER_ATOMIC);
  let code = 'sheetFeature(' + variable_name + ',"delete","row",' + value_num + ',' + value_delete_num + ');\n';
  return code;
};

Blockly.JavaScript['sheet_delete_column'] = function (block) {
  let variable_name = Blockly.JavaScript.variableDB_.getName(block.getFieldValue('name'), Blockly.Variables.NAME_TYPE);
  let value_num = Blockly.JavaScript.valueToCode(block, 'num', Blockly.JavaScript.ORDER_ATOMIC);
  let value_delete_num = Blockly.JavaScript.valueToCode(block, 'delete_num', Blockly.JavaScript.ORDER_ATOMIC);
  let code = 'sheetFeature(' + variable_name + ',"delete","column",' + value_num + ',' + value_delete_num + ');\n';
  return code;
};

Blockly.JavaScript['sheet_add_row'] = function (block) {
  let variable_name = Blockly.JavaScript.variableDB_.getName(block.getFieldValue('name'), Blockly.Variables.NAME_TYPE);
  let value_num = Blockly.JavaScript.valueToCode(block, 'num', Blockly.JavaScript.ORDER_ATOMIC);
  let value_add_num = Blockly.JavaScript.valueToCode(block, 'add_num', Blockly.JavaScript.ORDER_ATOMIC);
  let dropdown_add_type = block.getFieldValue('add_type');
  let code = 'sheetFeature(' + variable_name + ',"addRow","' + dropdown_add_type + '",' + value_num + ',' + value_add_num + ');\n';
  return code;
};

Blockly.JavaScript['sheet_add_column'] = function (block) {
  let variable_name = Blockly.JavaScript.variableDB_.getName(block.getFieldValue('name'), Blockly.Variables.NAME_TYPE);
  let value_num = Blockly.JavaScript.valueToCode(block, 'num', Blockly.JavaScript.ORDER_ATOMIC);
  let value_add_num = Blockly.JavaScript.valueToCode(block, 'add_num', Blockly.JavaScript.ORDER_ATOMIC);
  let dropdown_add_type = block.getFieldValue('add_type');
  let code = 'sheetFeature(' + variable_name + ',"addColumn","' + dropdown_add_type + '",' + value_num + ',' + value_add_num + ');\n';
  return code;
};