#include "pch.h"
using namespace Spire::Doc;

wstring GetHorizontalPositionType(HorizontalPosition value)
{
	switch (value)
	{
	case Spire::Doc::HorizontalPosition::None:
		//case Spire::Doc::HorizontalPosition::Default:
		return L"None";
		break;
	case Spire::Doc::HorizontalPosition::Left:
		return L"Left";
		break;
	case Spire::Doc::HorizontalPosition::Center:
		return L"Center";
		break;
	case Spire::Doc::HorizontalPosition::Right:
		return L"Right";
		break;
	case Spire::Doc::HorizontalPosition::Inside:
		return L"Inside";
		break;
	case Spire::Doc::HorizontalPosition::Outside:
		return L"Outside";
		break;
	case Spire::Doc::HorizontalPosition::Inline:
		return L"Inline";
		break;
	}
}
wstring GetHorizontalRelationType(HorizontalRelation value)
{
	switch (value)
	{
	case Spire::Doc::HorizontalRelation::Column:
		return L"Column";
		break;
	case Spire::Doc::HorizontalRelation::Margin:
		return L"Margin";
		break;
	case Spire::Doc::HorizontalRelation::Page:
		return L"Page";
		break;
	}
}
wstring GetVerticalPositionType(VerticalPosition value)
{
	switch (value)
	{
	case Spire::Doc::VerticalPosition::None:
		//case Spire::Doc::VerticalPosition::Default:
		return L"None";
		break;
	case Spire::Doc::VerticalPosition::Top:
		return L"Top";
		break;
	case Spire::Doc::VerticalPosition::Center:
		return L"Center";
		break;
	case Spire::Doc::VerticalPosition::Bottom:
		return L"Bottom";
		break;
	case Spire::Doc::VerticalPosition::Inside:
		return L"Inside";
		break;
	case Spire::Doc::VerticalPosition::Outside:
		return L"Outside";
		break;
	case Spire::Doc::VerticalPosition::Inline:
		return L"Inline";
		break;
	}
}
wstring GetVerticalRelationType(VerticalRelation value)
{
	switch (value)
	{
	case Spire::Doc::VerticalRelation::Margin:
		return L"Margin";
		break;
	case Spire::Doc::VerticalRelation::Page:
		return L"Page";
		break;
	case Spire::Doc::VerticalRelation::Paragraph:
		return L"Paragraph";
		break;
	}
}
int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableSample-Az.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetTablePosition.txt";

	//Create a document
	intrusive_ptr<Document> document = new Document();
	//Load file
	document->LoadFromFile(inputFile.c_str());
	//Get the first section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
	//Get the first table
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	wstring* stringBuidler = new wstring();

	//Verify whether the table uses "Around" text wrapping or not.
	if (table->GetTableFormat()->GetWrapTextAround())
	{
		intrusive_ptr<TablePositioning> positon = table->GetTableFormat()->GetPositioning();

		stringBuidler->append(L"Horizontal:");
		stringBuidler->append(L"Position:" + std::to_wstring(positon->GetHorizPosition()) + L" pt");
		stringBuidler->append(L"Absolute Position:" + GetHorizontalPositionType(positon->GetHorizPositionAbs()) + L", Relative to:" + GetHorizontalRelationType(positon->GetHorizRelationTo()));
		stringBuidler->append(L"");
		stringBuidler->append(L"Vertical:");
		stringBuidler->append(L"Position:" + std::to_wstring(positon->GetVertPosition()) + L" pt");
		stringBuidler->append(L"Absolute Position:" + GetVerticalPositionType(positon->GetVertPositionAbs()) + L", Relative to:" + GetVerticalRelationType(positon->GetVertRelationTo()));
		stringBuidler->append(L"");
		stringBuidler->append(L"Distance from surrounding text:");
		stringBuidler->append(L"Top:" + std::to_wstring(positon->GetDistanceFromTop()) + L" pt, Left:" + std::to_wstring(positon->GetDistanceFromLeft()) + L" pt");
		stringBuidler->append(L"Bottom:" + std::to_wstring(positon->GetDistanceFromBottom()) + L"pt, Right:" + std::to_wstring(positon->GetDistanceFromRight()) + L" pt");
	}

	//Save file.
	wofstream write(outputFile);
	write << stringBuidler->c_str();
	write.close();
	document->Close();
	delete stringBuidler;
}
