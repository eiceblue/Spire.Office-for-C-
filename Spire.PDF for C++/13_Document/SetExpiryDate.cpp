#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;


class StringBuilder
{
private:
	std::wstring privateString;

public:
	StringBuilder()
	{
	}

	StringBuilder(const std::wstring& initialString)
	{
		privateString = initialString;
	}
	StringBuilder* append(const std::wstring& toAppend)
	{
		privateString += toAppend;
		return this;
	}
	std::wstring toString()
	{
		return privateString;
	}
};

int main()
{
	wstring inputFile = DATAPATH"PDFTemplate-Az.pdf";
	wstring outputFile = OUTPUTPATH"SetExpiryDate.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	StringBuilder* sb = new StringBuilder();
	sb->append(L"var rightNow = new Date();");
	sb->append(L"var endDate = new Date('October 20, 2015 23:59:59');");
	sb->append(L"if(rightNow.getTime() > endDate)");
	sb->append(L"app.alert('This document has expired, please contact us for a new one.',1);");
	sb->append(L"this.closeDoc();");
	
	intrusive_ptr<PdfJavaScriptAction> js = new PdfJavaScriptAction(sb->toString().c_str());

	//doc->SetAfterOpenAction(Object::Convert<PdfAction>(js));
	doc->SetAfterOpenAction(js);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
