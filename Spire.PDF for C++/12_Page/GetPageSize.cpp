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

	StringBuilder* appendLine(const std::wstring& toAppend)
	{
		privateString += toAppend + L"\r\n";
		return this;
	}
	std::wstring toString()
	{
		return privateString;
	}
};


int main()
{
	wstring outputFile =OUTPUTPATH"GetPageSize.txt";
	wstring inputFile = DATAPATH"Sample.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first page of the loaded PDF file
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);

	//Get the width of page based on "point"
	float pointWidth = page->GetSize()->GetWidth();

	//Get the height of page
	float pointHeight = page->GetSize()->GetHeight();

	//Create PdfUnitConvertor to convert the unit
	intrusive_ptr<PdfUnitConvertor> unitCvtr = new PdfUnitConvertor();

	//Convert the size with "pixel"
	float pixelWidth = unitCvtr->ConvertUnits(pointWidth, PdfGraphicsUnit::Point, PdfGraphicsUnit::Pixel);
	float pixelHeight = unitCvtr->ConvertUnits(pointHeight, PdfGraphicsUnit::Point, PdfGraphicsUnit::Pixel);

	//Convert the size with "inch"
	float inchWidth = unitCvtr->ConvertUnits(pointWidth, PdfGraphicsUnit::Point, PdfGraphicsUnit::Inch);
	float inchHeight = unitCvtr->ConvertUnits(pointHeight, PdfGraphicsUnit::Point, PdfGraphicsUnit::Inch);

	//Convert the size with "centimeter"
	float centimeterWidth = unitCvtr->ConvertUnits(pointWidth, PdfGraphicsUnit::Point, PdfGraphicsUnit::Centimeter);
	float centimeterHeight = unitCvtr->ConvertUnits(pointHeight, PdfGraphicsUnit::Point, PdfGraphicsUnit::Centimeter);

	//Create StringBuilder to save 
	StringBuilder* content = new StringBuilder();

	//Add pointSize string to StringBuilder
	content->appendLine(L"The page size of the file is (width: " + std::to_wstring(pointWidth) + L"pt, height: " + std::to_wstring(pointHeight) + L"pt).");
	content->appendLine(L"The page size of the file is (width: " + std::to_wstring(pixelWidth) + L"pixel, height: " + std::to_wstring(pixelHeight) + L"pixel).");
	content->appendLine(L"The page size of the file is (width: " + std::to_wstring(inchWidth) + L"inch, height: " + std::to_wstring(inchHeight) + L"inch).");
	content->appendLine(L"The page size of the file is (width: " + std::to_wstring(centimeterWidth) + L"cm, height: " + std::to_wstring(centimeterHeight) + L"cm.)");

	//Save them to a txt file
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << content->toString();
	os.close();
}
