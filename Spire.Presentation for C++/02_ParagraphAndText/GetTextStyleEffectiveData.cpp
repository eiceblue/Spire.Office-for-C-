#include "pch.h"

using namespace Spire::Presentation;

namespace ParagraphAndText
{
	const wstring txtTypeToString(TextAlignmentType txtType)
	{
		switch (txtType)
		{
		case Spire::Presentation::TextAlignmentType::None:
			return L"None";
			break;
		case Spire::Presentation::TextAlignmentType::Left:
			return L"Left";
			break;
		case Spire::Presentation::TextAlignmentType::Center:
			return L"Center";
			break;
		case Spire::Presentation::TextAlignmentType::Right:
			return L"Right";
			break;
		case Spire::Presentation::TextAlignmentType::Justify:
			return L"Justify";
			break;
		case Spire::Presentation::TextAlignmentType::Dist:
			return L"Dist";
			break;
		default:
			break;
		}
	}
	const wstring triTypeToString(TriState  triType)
	{
		switch (triType)
		{
		case Spire::Presentation::TriState::Null:
			return L"Null";
			break;
		case Spire::Presentation::TriState::False:
			return L"False";
			break;
		case Spire::Presentation::TriState::True:
			return L"True";
			break;
		default:
			break;
		}
	}
}
	const wstring fontTypeToString(FontAlignmentType fontType)
	{
		switch (fontType)
		{
		case Spire::Presentation::FontAlignmentType::None:
			return L"None";
			break;
		case Spire::Presentation::FontAlignmentType::Auto:
			return L"Auto";
			break;
		case Spire::Presentation::FontAlignmentType::Top:
			return L"Top";
			break;
		case Spire::Presentation::FontAlignmentType::Center:
			return L"Center";
			break;
		case Spire::Presentation::FontAlignmentType::Bottom:
			return L"Bottom";
			break;
		case Spire::Presentation::FontAlignmentType::Baseline:
			return L"Baseline";
			break;
		default:
			break;
		}
	}


int main()
{
	
	wstring inputFile = DATAPATH"Template_Az1.pptx";
	wstring outputFile = OUTPUTPATH"GetTextStyleEffectiveData.txt";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	wofstream outFile(outputFile, ios::out);

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);
	//Get a shape 
	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	for (int p = 0; p < shape->GetTextFrame()->GetParagraphs()->GetCount(); p++)
	{
		auto paragraph = shape->GetTextFrame()->GetParagraphs()->GetItem(p);
		outFile << "Text style for Paragraph " << std::to_string(p).c_str() << ":" << endl;
		//Get the paragraph style
		outFile << " Indent: " << paragraph->GetIndent() << endl;
		outFile << " Alignment: " << ParagraphAndText::txtTypeToString(paragraph->GetAlignment()) << endl;
		outFile << " Font alignment: " << fontTypeToString(paragraph->GetFontAlignment()) << endl;
		outFile << " Hanging punctuation: " << ParagraphAndText::triTypeToString(paragraph->GetHangingPunctuation()) << endl;
		outFile << " Line spacing: " << paragraph->GetLineSpacing() << endl;
		outFile << " Space before: " << paragraph->GetSpaceBefore() << endl;
		outFile << " Space after: " << paragraph->GetSpaceAfter() << endl;
		outFile << endl;
		for (int r = 0; r < paragraph->GetTextRanges()->GetCount(); r++)
		{
			auto textRange = paragraph->GetTextRanges()->GetItem(r);
			outFile << "  Text style for Paragraph " << std::to_string(p).c_str() << " TextRange " << std::to_string(r).c_str() << " :" << endl;
			//Get the text range style
			outFile << "    Font height: " << textRange->GetFontHeight() << endl;
			outFile << "    Language: " << textRange->GetLanguage() << endl;
			outFile << "    Font: " << textRange->GetLatinFont()->GetFontName() << endl;
			outFile << endl;
		}
	}
	outFile.close();
	

}

