#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"AddHyperlinkToText.pptx";
	wstring outputFile = OUTPUTPATH"AddHyperlinkToText.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Find the text we want to add link to it.
	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	intrusive_ptr<TextParagraph> tp = shape->GetTextFrame()->GetTextRange()->GetParagraph();
	std::wstring temp = tp->GetText();

	//Split the original text.
	std::wstring textToLink = L"Spire.Presentation";
	std::wstring::size_type pos = temp.find(textToLink);
	std::wstring strSplit = temp.substr(0, pos);

	//Clear all text.
	tp->GetTextRanges()->Clear();

	//Add new text.
	intrusive_ptr<TextRange> tr = new TextRange(strSplit.c_str());
	tp->GetTextRanges()->Append(tr);

	//Add the hyperlink.
	tr = new TextRange(textToLink.c_str());
	tr->GetClickAction()->SetAddress(L"https://www.e-iceblue.com/Introduce/presentation-for-CPP.html");
	tp->GetTextRanges()->Append(tr);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
