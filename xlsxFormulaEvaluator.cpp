#include "stdafx.h"
#include"xlsxFormulaEvaluator.h"



TreeFormulaEvaluator::TreeFormulaEvaluator(Book* b) : book(b) {
    // Cache all sheets
    for (int i = 0; i < book->sheetCount(); ++i) {
        Sheet* sheet = book->getSheet(i);
        if (sheet) {
            std::wstring wname = sheet->name();
            std::string name(wname.begin(), wname.end());
            sheets[name] = sheet;
        }
    }
}

// Tokenize formula string
std::vector<Token> TreeFormulaEvaluator::tokenize(const std::string& formula) {
    std::vector<Token> tokens;
    std::string cleanFormula = formula;

    // Remove = and leading +
    if (!cleanFormula.empty() && cleanFormula[0] == '=') {
        cleanFormula = cleanFormula.substr(1);
    }
    if (!cleanFormula.empty() && cleanFormula[0] == '+') {
        cleanFormula = cleanFormula.substr(1);
    }

    size_t i = 0;
    while (i < cleanFormula.length()) {
        char c = cleanFormula[i];

        // Skip whitespace
        if (std::isspace(c)) {
            i++;
            continue;
        }

        // Operators and parentheses
        if (c == '+' || c == '-' || c == '*' || c == '/') {
            tokens.emplace_back(Token::OPERATOR, std::string(1, c));
            i++;
        }
        else if (c == '(') {
            tokens.emplace_back(Token::LPAREN, "(");
            i++;
        }
        else if (c == ')') {
            tokens.emplace_back(Token::RPAREN, ")");
            i++;
        }
        else if (c == ',') {
            tokens.emplace_back(Token::COMMA, ",");
            i++;
        }
        // Numbers (constants)
        else if (std::isdigit(c) || c == '.') {
            std::string number;
            while (i < cleanFormula.length() && (std::isdigit(cleanFormula[i]) || cleanFormula[i] == '.')) {
                number += cleanFormula[i++];
            }
            tokens.emplace_back(Token::CONSTANT, number);
        }
        // Functions, cell references, ranges
        else if (std::isalpha(c)) {
            std::string identifier;
            std::string sheetName;

            // Read the identifier
            while (i < cleanFormula.length() &&
                (std::isalnum(cleanFormula[i]) || cleanFormula[i] == '_')) {
                identifier += cleanFormula[i++];
            }

            // Check for sheet reference (!)
            if (i < cleanFormula.length() && cleanFormula[i] == '!') {
                sheetName = identifier;
                identifier = "";
                i++; // skip !

                // Read the actual cell reference after !
                while (i < cleanFormula.length() &&
                    (std::isalnum(cleanFormula[i]) || cleanFormula[i] == '_')) {
                    identifier += cleanFormula[i++];
                }
            }

            // Check if it's followed by ( - then it's a function
            if (i < cleanFormula.length() && cleanFormula[i] == '(') {
                tokens.emplace_back(Token::FUNCTION, identifier);
            }
            // Check if it contains : - then it might be a range
            else if (i < cleanFormula.length() && cleanFormula[i] == ':') {
                // This is start of a range, read the full range
                std::string range = identifier;
                range += cleanFormula[i++]; // add :

                // Read end of range
                while (i < cleanFormula.length() &&
                    (std::isalnum(cleanFormula[i]) || cleanFormula[i] == '_')) {
                    range += cleanFormula[i++];
                }
                tokens.emplace_back(Token::RANGE, range, sheetName);
            }
            // Regular cell reference
            else {
                tokens.emplace_back(Token::CELL_REF, identifier, sheetName);
            }
        }
        else {
            // Unknown character, skip
            i++;
        }
    }

    return tokens;
}

// Parse tokens into expression tree
std::shared_ptr<FormulaNode> TreeFormulaEvaluator::parse(const std::vector<Token>& tokens) {
    size_t index = 0;
    return parseExpression(tokens, index);
}

// Parse expression with operator precedence
std::shared_ptr<FormulaNode> TreeFormulaEvaluator::parseExpression(const std::vector<Token>& tokens, size_t& index) {
    return parseAddition(tokens, index);
}

std::shared_ptr<FormulaNode> TreeFormulaEvaluator::parseAddition(const std::vector<Token>& tokens, size_t& index) {
    auto left = parseMultiplication(tokens, index);

    while (index < tokens.size() &&
        tokens[index].type == Token::OPERATOR &&
        (tokens[index].value == "+" || tokens[index].value == "-")) {

        std::string op = tokens[index].value;
        index++;
        auto right = parseMultiplication(tokens, index);

        auto opNode = std::make_shared<FormulaNode>(FormulaNode::OPERATOR, op);
        opNode->addChild(left);
        opNode->addChild(right);
        left = opNode;
    }

    return left;
}

std::shared_ptr<FormulaNode> TreeFormulaEvaluator::parseMultiplication(const std::vector<Token>& tokens, size_t& index) {
    auto left = parseFactor(tokens, index);

    while (index < tokens.size() &&
        tokens[index].type == Token::OPERATOR &&
        (tokens[index].value == "*" || tokens[index].value == "/")) {

        std::string op = tokens[index].value;
        index++;
        auto right = parseFactor(tokens, index);

        auto opNode = std::make_shared<FormulaNode>(FormulaNode::OPERATOR, op);
        opNode->addChild(left);
        opNode->addChild(right);
        left = opNode;
    }

    return left;
}

std::shared_ptr<FormulaNode> TreeFormulaEvaluator::parseFactor(const std::vector<Token>& tokens, size_t& index) {
    if (index >= tokens.size()) return nullptr;

    const Token& token = tokens[index];

    // Parentheses
    if (token.type == Token::LPAREN) {
        index++; // skip (
        auto expr = parseExpression(tokens, index);
        if (index < tokens.size() && tokens[index].type == Token::RPAREN) {
            index++; // skip )
        }
        return expr;
    }

    // Functions
    if (token.type == Token::FUNCTION) {
        std::string funcName = token.value;
        index++; // skip function name

        auto funcNode = std::make_shared<FormulaNode>(FormulaNode::FUNCTION, funcName);

        if (index < tokens.size() && tokens[index].type == Token::LPAREN) {
            index++; // skip (

            // Parse function arguments
            while (index < tokens.size() && tokens[index].type != Token::RPAREN) {
                if (tokens[index].type == Token::RANGE) {
                    auto rangeNode = std::make_shared<FormulaNode>(
                        FormulaNode::RANGE,
                        tokens[index].value,
                        tokens[index].sheetName
                    );
                    funcNode->addChild(rangeNode);
                    index++;
                }
                else if (tokens[index].type == Token::COMMA) {
                    index++; // skip comma
                }
                else {
                    auto arg = parseExpression(tokens, index);
                    if (arg) funcNode->addChild(arg);
                }
            }

            if (index < tokens.size() && tokens[index].type == Token::RPAREN) {
                index++; // skip )
            }
        }

        return funcNode;
    }

    // Cell references
    if (token.type == Token::CELL_REF) {
        auto cellNode = std::make_shared<FormulaNode>(
            FormulaNode::CELL_REF,
            token.value,
            token.sheetName
        );
        index++;
        return cellNode;
    }

    // Constants
    if (token.type == Token::CONSTANT) {
        auto constNode = std::make_shared<FormulaNode>(FormulaNode::CONSTANT, token.value);
        index++;
        return constNode;
    }

    // Ranges (standalone)
    if (token.type == Token::RANGE) {
        auto rangeNode = std::make_shared<FormulaNode>(
            FormulaNode::RANGE,
            token.value,
            token.sheetName
        );
        index++;
        return rangeNode;
    }

    return nullptr;
}

double TreeFormulaEvaluator::evaluateFunction(std::shared_ptr<FormulaNode> node) {
    if (node->value == "SUM") {
        return evaluateSum(node);
    }
    // Add other functions as needed (AVERAGE, MAX, etc.)
    return 0.0;
}

// Evaluate the tree
double TreeFormulaEvaluator::evaluate(std::shared_ptr<FormulaNode> node)  {
    if (!node) return 0.0;

    // Check for circular reference
    if (node->isEvaluating) {
        std::cout << "Circular reference detected!" << std::endl;
        return 0.0;
    }

    // Return cached result if already evaluated
    if (node->isEvaluated) {
        return node->cachedResult;
    }

    node->isEvaluating = true;
    double result = 0.0;

    switch (node->type) {
    case FormulaNode::CONSTANT:
        result = std::stod(node->value);
        break;

    case FormulaNode::CELL_REF:
        result = evaluateCellReference(node);
        break;

    case FormulaNode::FUNCTION:
        result = evaluateFunction(node);
        break;

    case FormulaNode::OPERATOR:
        result = evaluateOperator(node);
        break;

    case FormulaNode::RANGE:
        // Ranges are typically handled within functions
        result = 0.0;
        break;
    }

    node->isEvaluating = false;
    node->isEvaluated = true;
    node->cachedResult = result;

    return result;
}

double TreeFormulaEvaluator::evaluateCellReference(std::shared_ptr<FormulaNode> node)  {
    Sheet* sheet = getSheet(node->sheetName);
    if (!sheet) return 0.0;

    auto [row, col] = parseCellAddress(node->value);
    if (row < 0 || col < 0) return 0.0;

    std::cout << "Evaluating cell " << node->toString() << std::endl;

    CellType cellType = sheet->cellType(row, col);
	bool isFormula = sheet->isFormula(row, col);
    if (!isFormula && (cellType == CELLTYPE_STRING || cellType == CELLTYPE_NUMBER)) {
        if(cellType == CELLTYPE_NUMBER) {
            // Direct numeric value
            double value = sheet->readNum(row, col);
            std::cout << "  Direct numeric value: " << value << std::endl;
            return value;
		}
        std::wstring cvalue = sheet->readStr(row, col);
        std::string svalue(cvalue.begin(), cvalue.end());
        double value = std::stoi(svalue);
        //double value = sheet->readNum(row, col);
        std::cout << "  Direct value: " << value << std::endl;
        return value;
    }
    else if (isFormula/*cellType == CELLTYPE_FORMULA*/) {
        // Try pre-calculated value first
        double preCalc = sheet->readNum(row, col);
        if (preCalc != 0.0) {
            std::cout << "  Pre-calculated: " << preCalc << std::endl;
            return preCalc;
        }

        // Evaluate formula recursively
        const wchar_t* formula = sheet->readFormula(row, col);
        if (formula) {
            std::string formulaStr(formula, formula + wcslen(formula));
            std::wcout << L"  Has formula: " << formula << std::endl;

            auto tokens = tokenize(formulaStr);
            auto tree = parse(tokens);
            double result = evaluate(tree);
            std::cout << "  Recursive result: " << result << std::endl;
            return result;
        }
    }


    return 0.0;
}

double TreeFormulaEvaluator::evaluateSum(std::shared_ptr<FormulaNode> node) {
    double sum = 0.0;
    std::cout << "Evaluating SUM function" << std::endl;

    for (auto& child : node->children) {
        if (child->type == FormulaNode::RANGE) {
            sum += evaluateRange(child);
        }
        else {
            sum += evaluate(child);
        }
    }

    std::cout << "SUM result: " << sum << std::endl;
    return sum;
}

double TreeFormulaEvaluator::evaluateRange(std::shared_ptr<FormulaNode> node) {
    std::string range = node->value;
    std::cout << "Evaluating range: " << node->toString() << std::endl;

    // Parse range like "G22:L22"
    size_t colonPos = range.find(':');
    if (colonPos == std::string::npos) {
        // Single cell
        auto singleCell = std::make_shared<FormulaNode>(
            FormulaNode::CELL_REF, range, node->sheetName);
        return evaluate(singleCell);
    }

    std::string startCell = range.substr(0, colonPos);
    std::string endCell = range.substr(colonPos + 1);

    auto [startRow, startCol] = parseCellAddress(startCell);
    auto [endRow, endCol] = parseCellAddress(endCell);

    Sheet* sheet = getSheet(node->sheetName);
    if (!sheet) return 0.0;

    double sum = 0.0;
    for (int row = startRow; row <= endRow; ++row) {
        for (int col = startCol; col <= endCol; ++col) {
            std::string cellAddr = columnToLetter(col) + std::to_string(row + 1);
            auto cellNode = std::make_shared<FormulaNode>(
                FormulaNode::CELL_REF, cellAddr, node->sheetName);
            sum += evaluate(cellNode);
        }
    }

    return sum;
}

double TreeFormulaEvaluator::evaluateOperator(std::shared_ptr<FormulaNode> node) {
    if (node->children.size() != 2) return 0.0;

    double left = evaluate(node->children[0]);
    double right = evaluate(node->children[1]);

    std::cout << "Operator " << node->value << ": " << left << " " << node->value << " " << right;

    if (node->value == "+") {
        double result = left + right;
        std::cout << " = " << result << std::endl;
        return result;
    }
    else if (node->value == "-") {
        double result = left - right;
        std::cout << " = " << result << std::endl;
        return result;
    }
    else if (node->value == "*") {
        double result = left * right;
        std::cout << " = " << result << std::endl;
        return result;
    }
    else if (node->value == "/") {
        double result = (right != 0) ? left / right : 0.0;
        std::cout << " = " << result << std::endl;
        return result;
    }

    return 0.0;
}

Sheet* TreeFormulaEvaluator::getSheet(const std::string& sheetName) {
    if (sheetName.empty()) {
        return book->getSheet(0); // Default sheet
    }

    auto it = sheets.find(sheetName);
    return (it != sheets.end()) ? it->second : nullptr;
}

std::pair<int, int> TreeFormulaEvaluator::parseCellAddress(const std::string& cellAddr) {
    std::string colPart, rowPart;

    for (char c : cellAddr) {
        if (std::isalpha(c)) {
            colPart += c;
        }
        else if (std::isdigit(c)) {
            rowPart += c;
        }
    }

    if (colPart.empty() || rowPart.empty()) {
        return { -1, -1 };
    }

    int col = 0;
    for (char c : colPart) {
        col = col * 26 + (std::toupper(c) - 'A' + 1);
    }
    col--; // Convert to 0-based

    int row = std::stoi(rowPart) - 1; // Convert to 0-based

    return { row, col };
}

std::string TreeFormulaEvaluator::columnToLetter(int col) {
    std::string result;
    while (col >= 0) {
        result = char('A' + (col % 26)) + result;
        col = col / 26 - 1;
    }
    return result;
}

// Main evaluation function
double TreeFormulaEvaluator::evaluateFormula(const std::string& formula) {
    std::cout << "\n=== Evaluating Formula: " << formula << " ===" << std::endl;

    auto tokens = tokenize(formula);

    std::cout << "Tokens: ";
    for (const auto& token : tokens) {
        std::cout << "[" << token.value << "] ";
    }
    std::cout << std::endl;

    auto tree = parse(tokens);
    if (!tree) {
        std::cout << "Failed to parse formula!" << std::endl;
        return 0.0;
    }

    double result = evaluate(tree);
    std::cout << "Final result: " << result << std::endl;
    return result;
}

// Helper to print tree structure (for debugging puspose only )
void TreeFormulaEvaluator::printTree(std::shared_ptr<FormulaNode> node, int depth)  {
    if (!node) return;

    std::string indent(depth * 2, ' ');
    std::cout << indent << node->toString() << " (" <<
        (node->type == FormulaNode::OPERATOR ? "OP" :
            node->type == FormulaNode::FUNCTION ? "FUNC" :
            node->type == FormulaNode::CELL_REF ? "CELL" :
            node->type == FormulaNode::CONSTANT ? "CONST" : "RANGE")
        << ")" << std::endl;

    for (auto& child : node->children) {
        printTree(child, depth + 1);
    }
}
