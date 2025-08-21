
#include "xlsxFormulaEvaluator.h"

int main(){
Book* pBook = xlCreateBook();

CString sourceFile = L"C:\\T1\\346043(11_106).xls";

//pBook->setKey(_T(""), <your_key>);
bool isFileLoaded = pBook->load(sourceFile);
TreeFormulaEvaluator evaluator(pBook);


// Test formulae
std::vector<std::string> testFormulas = {
				//"=U25/SUM(G22:L22)*0.9144" -- tested working fine
				//"=U25/SUM(G22:L22)"  -- tested working fine
				// "=W25/SUM(G22:L22)" -- tsted working fine 
				//"=W25/SUM(G22:L22)*12" -- tested working fine
				//"=(+G23*APPENDIX!C4+H23*APPENDIX!D4+I23*APPENDIX!E4+J23*APPENDIX!F4+K23*APPENDIX!G4+L23*APPENDIX!H4)/U25", -- tsted works fine						
				// "=I11-H11" -- tested working fine
				// "=U25" -- tested working fine
};

for (const auto& formula : testFormulas) {
				double result = evaluator.evaluateFormula(formula);
				std::cout << "Formula: " << result << std::endl;
				std::cout << "================================\n" << std::endl;
}				

pBook->release();

return 0;
}