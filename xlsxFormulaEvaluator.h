#pragma once
#include "stdafx.h"
#include "libxl.h"
#include <iostream>
#include <string>
#include <vector>
#include <sstream>
#include <algorithm>
#include <cctype>
#include <map>
#include <regex>
#include <cmath>
#include <set>
#include <memory>
#include "exprtk.hpp"

using namespace libxl;
// Token structure for parsing
struct Token {
    enum Type {
        CELL_REF,     // G23, APPENDIX!C4
        FUNCTION,     // SUM, AVERAGE
        OPERATOR,     // +, -, *, /
        CONSTANT,     // 12, 0.9144
        LPAREN,       // (
        RPAREN,       // )
        RANGE,        // G22:L22
        COMMA,         // ,
        EXCLAMATION, // !

    };

    Type type;
    std::string value;
    std::string sheetName; // for cross-sheet references

    Token(Type t, const std::string& v, const std::string& sheet = "")
        : type(t), value(v), sheetName(sheet) {
    }
};

// Formula tree node
class FormulaNode {
public:
    enum Type {
        CELL_REF,
        FUNCTION,
        OPERATOR,
        CONSTANT,
        RANGE
    };

    Type type;
    std::string value;
    std::string sheetName;
    std::vector<std::shared_ptr<FormulaNode>> children;

    // Evaluation state
    mutable double cachedResult = 0.0;
    mutable bool isEvaluated = false;
    mutable bool isEvaluating = false;

    FormulaNode(Type t, const std::string& v, const std::string& sheet = "")
        : type(t), value(v), sheetName(sheet) {
    }

    void addChild(std::shared_ptr<FormulaNode> child) {
        children.push_back(child);
    }

    std::string toString() const {
        std::string result;
        if (!sheetName.empty()) result += sheetName + "!";
        result += value;
        return result;
    }
};

class TreeFormulaEvaluator {
private:
    Book* book;
    std::map<std::string, Sheet*> sheets;

public:
    TreeFormulaEvaluator(Book* b);

    // Tokenize formula string
    std::vector<Token> tokenize(const std::string& formula);

    // Parse tokens into expression tree
    std::shared_ptr<FormulaNode> parse(const std::vector<Token>& tokens);

private:
    // Parse expression with operator precedence
    std::shared_ptr<FormulaNode> parseExpression(const std::vector<Token>& tokens, size_t& index);

    std::shared_ptr<FormulaNode> parseAddition(const std::vector<Token>& tokens, size_t& index);

    std::shared_ptr<FormulaNode> parseMultiplication(const std::vector<Token>& tokens, size_t& index);

    std::shared_ptr<FormulaNode> parseFactor(const std::vector<Token>& tokens, size_t& index);

public:
    // Evaluate the tree
    double evaluate(std::shared_ptr<FormulaNode> node);

private:
    double evaluateCellReference(std::shared_ptr<FormulaNode> node);

    double evaluateFunction(std::shared_ptr<FormulaNode> node);

    double evaluateSum(std::shared_ptr<FormulaNode> node);

    double evaluateRange(std::shared_ptr<FormulaNode> node);

    double evaluateOperator(std::shared_ptr<FormulaNode> node);

    Sheet* getSheet(const std::string& sheetName);

    std::pair<int, int> parseCellAddress(const std::string& cellAddr);

    std::string columnToLetter(int col);

public:
    // Main evaluation function
    double evaluateFormula(const std::string& formula);

    // Helper to print tree structure (for debugging)
    void printTree(std::shared_ptr<FormulaNode> node, int depth = 0);
};
