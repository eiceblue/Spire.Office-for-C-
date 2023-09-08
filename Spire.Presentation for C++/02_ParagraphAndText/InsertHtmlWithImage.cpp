#include "pch.h"

using namespace Spire::Presentation;

int main()
{

	wstring outputFile = OUTPUTPATH"InsertHtmlWithImage.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();
	intrusive_ptr<ShapeList> shapes = ppt->GetSlides()->GetItem(0)->GetShapes();

	shapes->AddFromHtml(L"<html><div><p>First paragraph</p><p><img src='../../../Data/TestData/Demo/Logo.png'/></p><p>Second paragraph </p></html>");

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}