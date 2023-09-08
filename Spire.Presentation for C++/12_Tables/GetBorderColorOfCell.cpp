#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"GetBorderColorOfCell.pptx";
	wstring outputFile = OUTPUTPATH"GetBorderColorOfCell.txt";

	DeleteFile(outputFile.c_str());
	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());
	//Get the table in the first slide
	intrusive_ptr<ITable> table = Object::Dynamic_cast<ITable>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Get borders' color of the first cell
	wofstream outFile(outputFile);
	outFile << "Color of left border:Color [A=" << table->GetItem(0, 0)->GetBorderLeftDisplayColor()->GetA() << L",R =" << table->GetItem(0, 0)->GetBorderLeftDisplayColor()->GetR() << L",G =" << table->GetItem(0, 0)->GetBorderLeftDisplayColor()->GetG() << L",B =" << table->GetItem(0, 0)->GetBorderLeftDisplayColor()->GetB() << L"]" << endl;
	outFile << "Color of top border:Color [A=" << table->GetItem(0, 0)->GetBorderTopDisplayColor()->GetA() << L",R =" << table->GetItem(0, 0)->GetBorderTopDisplayColor()->GetR() << L",G =" << table->GetItem(0, 0)->GetBorderTopDisplayColor()->GetG() << L",B =" << table->GetItem(0, 0)->GetBorderTopDisplayColor()->GetB() << L"]" << endl;
	outFile << "Color of right border:Color [A=" << table->GetItem(0, 0)->GetBorderRightDisplayColor()->GetA() << L",R =" << table->GetItem(0, 0)->GetBorderRightDisplayColor()->GetR() << L",G =" << table->GetItem(0, 0)->GetBorderRightDisplayColor()->GetG() << L",B =" << table->GetItem(0, 0)->GetBorderRightDisplayColor()->GetB() << L"]" << endl;
	outFile << "Color of bottom border:Color [A=" << table->GetItem(0, 0)->GetBorderBottomDisplayColor()->GetA() << L",R =" << table->GetItem(0, 0)->GetBorderBottomDisplayColor()->GetR() << L",G =" << table->GetItem(0, 0)->GetBorderBottomDisplayColor()->GetG() << L",B =" << table->GetItem(0, 0)->GetBorderBottomDisplayColor()->GetB() << L"]" << endl;

	//Get display color of the first cell
	outFile << "Color of cell:Color [A=" << table->GetItem(0, 0)->GetDisplayColor()-> GetA() << L",R =" << table->GetItem(0, 0)->GetDisplayColor()->GetR() << L",G =" << table->GetItem(0, 0)->GetDisplayColor()->GetG() << L",B =" << table->GetItem(0, 0)->GetDisplayColor()->GetB() << L"]" << endl;
	outFile.close();
		
}
