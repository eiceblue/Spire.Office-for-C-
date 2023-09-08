#include "pch.h"
using namespace Spire::Presentation;
namespace SmartArt
{
	const wstring LayoutTypeToString(SmartArtLayoutType layoutType)
	{
		switch (layoutType)
		{
		case Spire::Presentation::SmartArtLayoutType::BasicBlockList:
			return L"BasicBlockList";
			break;
		case Spire::Presentation::SmartArtLayoutType::PictureCaptionList:
			return L"PictureCaptionList";
			break;
		case Spire::Presentation::SmartArtLayoutType::VerticalBulletList:
			return L"VerticalBulletList";
			break;
		case Spire::Presentation::SmartArtLayoutType::VerticalBoxList:
			return L"VerticalBoxList";
			break;
		case Spire::Presentation::SmartArtLayoutType::HorizontalBulletList:
			return L"HorizontalBulletList";
			break;
		case Spire::Presentation::SmartArtLayoutType::PictureAccentList:
			return L"PictureAccentList";
			break;
		case Spire::Presentation::SmartArtLayoutType::BendingPictureAccentList:
			return L"BendingPictureAccentList";
			break;
		case Spire::Presentation::SmartArtLayoutType::StackedList:
			return L"StackedList";
			break;
		case Spire::Presentation::SmartArtLayoutType::DetailedProcess:
			return L"DetailedProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::GroupedList:
			return L"GroupedList";
			break;
		case Spire::Presentation::SmartArtLayoutType::HorizontalPictureList:
			return L"HorizontalPictureList";
			break;
		case Spire::Presentation::SmartArtLayoutType::ContinuousPictureList:
			return L"ContinuousPictureList";
			break;
		case Spire::Presentation::SmartArtLayoutType::VerticalPictureList:
			return L"VerticalPictureList";
			break;
		case Spire::Presentation::SmartArtLayoutType::VerticalPictureAccentList:
			return L"VerticalPictureAccentList";
			break;
		case Spire::Presentation::SmartArtLayoutType::VerticalBlockList:
			return L"VerticalBlockList";
			break;
		case Spire::Presentation::SmartArtLayoutType::VerticalChevronList:
			return L"VerticalChevronList";
			break;
		case Spire::Presentation::SmartArtLayoutType::VerticalArrowList:
			return L"VerticalArrowList";
			break;
		case Spire::Presentation::SmartArtLayoutType::TrapezoidList:
			return L"TrapezoidList";
			break;
		case Spire::Presentation::SmartArtLayoutType::TableList:
			return L"TableList";
			break;
		case Spire::Presentation::SmartArtLayoutType::PyramidList:
			return L"PyramidList";
			break;
		case Spire::Presentation::SmartArtLayoutType::TargetList:
			return L"TargetList";
			break;
		case Spire::Presentation::SmartArtLayoutType::HierarchyList:
			return L"HierarchyList";
			break;
		case Spire::Presentation::SmartArtLayoutType::TableHierarhy:
			return L"TableHierarhy";
			break;
		case Spire::Presentation::SmartArtLayoutType::BasicProcess:
			return L"BasicProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::AccentProcess:
			return L"AccentProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::PictureAccentProcess:
			return L"PictureAccentProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::AlternatingFlow:
			return L"AlternatingFlow";
			break;
		case Spire::Presentation::SmartArtLayoutType::ContinuousBlockProcess:
			return L"ContinuousBlockProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::ContinuousArrowProcess:
			return L"ContinuousArrowProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::ProcessArrows:
			return L"ProcessArrows";
			break;
		case Spire::Presentation::SmartArtLayoutType::BasicTimeline:
			return L"BasicTimeline";
			break;
		case Spire::Presentation::SmartArtLayoutType::BasicChevronProcess:
			return L"BasicChevronProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::ClosedChevronProcess:
			return L"ClosedChevronProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::ChevronList:
			return L"ChevronList";
			break;
		case Spire::Presentation::SmartArtLayoutType::VerticalProcess:
			return L"VerticalProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::StaggeredProcess:
			return L"StaggeredProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::ProcessList:
			return L"ProcessList";
			break;
		case Spire::Presentation::SmartArtLayoutType::SegmentedProcess:
			return L"SegmentedProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::BasicBendingProcess:
			return L"BasicBendingProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::RepeatingBendingProcess:
			return L"RepeatingBendingProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::VerticalBendingProcess:
			return L"VerticalBendingProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::UpwardArrow:
			return L"UpwardArrow";
			break;
		case Spire::Presentation::SmartArtLayoutType::CircularBendingProcess:
			return L"CircularBendingProcess";
			break;
		case Spire::Presentation::SmartArtLayoutType::Equation:
			return L"Equation";
			break;
		case Spire::Presentation::SmartArtLayoutType::VerticalEquation:
			return L"VerticalEquation";
			break;
		case Spire::Presentation::SmartArtLayoutType::Funnel:
			return L"Funnel";
			break;
		case Spire::Presentation::SmartArtLayoutType::Gear:
			return L"Gear";
			break;
		case Spire::Presentation::SmartArtLayoutType::ArrowRibbon:
			return L"ArrowRibbon";
			break;
		case Spire::Presentation::SmartArtLayoutType::OpposingArrows:
			return L"OpposingArrows";
			break;
		case Spire::Presentation::SmartArtLayoutType::ConvergingArrows:
			return L"ConvergingArrows";
			break;
		case Spire::Presentation::SmartArtLayoutType::DivergingArrows:
			return L"DivergingArrows";
			break;
		case Spire::Presentation::SmartArtLayoutType::BasicCycle:
			return L"BasicCycle";
			break;
		case Spire::Presentation::SmartArtLayoutType::TextCycle:
			return L"TextCycle";
			break;
		case Spire::Presentation::SmartArtLayoutType::BlockCycle:
			return L"BlockCycle";
			break;
		case Spire::Presentation::SmartArtLayoutType::NondirectionalCycle:
			return L"NondirectionalCycle";
			break;
		case Spire::Presentation::SmartArtLayoutType::ContinuousCycle:
			return L"ContinuousCycle";
			break;
		case Spire::Presentation::SmartArtLayoutType::MultidirectionalCycle:
			return L"MultidirectionalCycle";
			break;
		case Spire::Presentation::SmartArtLayoutType::SegmentedCycle:
			return L"SegmentedCycle";
			break;
		case Spire::Presentation::SmartArtLayoutType::BasicPie:
			return L"BasicPie";
			break;
		case Spire::Presentation::SmartArtLayoutType::RadialCycle:
			return L"RadialCycle";
			break;
		case Spire::Presentation::SmartArtLayoutType::BasicRadial:
			return L"BasicRadial";
			break;
		case Spire::Presentation::SmartArtLayoutType::DivergingRadial:
			return L"DivergingRadial";
			break;
		case Spire::Presentation::SmartArtLayoutType::RadialVenn:
			return L"RadialVenn";
			break;
		case Spire::Presentation::SmartArtLayoutType::CycleMatrix:
			return L"CycleMatrix";
			break;
		case Spire::Presentation::SmartArtLayoutType::OrganizationChart:
			return L"OrganizationChart";
			break;
		case Spire::Presentation::SmartArtLayoutType::Hierarchy:
			return L"Hierarchy";
			break;
		case Spire::Presentation::SmartArtLayoutType::LabeledHierarhy:
			return L"LabeledHierarhy";
			break;
		case Spire::Presentation::SmartArtLayoutType::HorizontalHierarhy:
			return L"HorizontalHierarhy";
			break;
		case Spire::Presentation::SmartArtLayoutType::HorizontalLabeledHierarhy:
			return L"HorizontalLabeledHierarhy";
			break;
		case Spire::Presentation::SmartArtLayoutType::Balance:
			return L"Balance";
			break;
		case Spire::Presentation::SmartArtLayoutType::CounterbalanceArrow:
			return L"CounterbalanceArrow";
			break;
		case Spire::Presentation::SmartArtLayoutType::SegmentedPyramid:
			return L"SegmentedPyramid";
			break;
		case Spire::Presentation::SmartArtLayoutType::NestedTarget:
			return L"NestedTarget";
			break;
		case Spire::Presentation::SmartArtLayoutType::ConvergingRadial:
			return L"ConvergingRadial";
			break;
		case Spire::Presentation::SmartArtLayoutType::RadialList:
			return L"RadialList";
			break;
		case Spire::Presentation::SmartArtLayoutType::BasicTarget:
			return L"BasicTarget";
			break;
		case Spire::Presentation::SmartArtLayoutType::BasicVenn:
			return L"BasicVenn";
			break;
		case Spire::Presentation::SmartArtLayoutType::LinearVenn:
			return L"LinearVenn";
			break;
		case Spire::Presentation::SmartArtLayoutType::StacketVenn:
			return L"StacketVenn";
			break;
		case Spire::Presentation::SmartArtLayoutType::BasicMatrix:
			return L"BasicMatrix";
			break;
		case Spire::Presentation::SmartArtLayoutType::TitledMatrix:
			return L"TitledMatrix";
			break;
		case Spire::Presentation::SmartArtLayoutType::GridMatrix:
			return L"GridMatrix";
			break;
		case Spire::Presentation::SmartArtLayoutType::BasicPyramid:
			return L"BasicPyramid";
			break;
		case Spire::Presentation::SmartArtLayoutType::InvertedPyramid:
			return L"InvertedPyramid";
			break;
		default:
			break;
		}
	}
}
int main() {
	wstring inputFile = DATAPATH"SmartArt.pptx";
	wstring outputFile = OUTPUTPATH"AccessSmartArtLayout.txt";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	wstring* content = new std::wstring();

	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		//Get the SmartArt and collect nodes
		intrusive_ptr<ISmartArt> sa = Object::Dynamic_cast<ISmartArt>(shape);
		if (sa != nullptr)
		{
			//Check SmartArt Layout	
			SmartArtLayoutType layout = sa->GetLayoutType();
			content->append(L"SmartArt layout type is " + SmartArt::LayoutTypeToString(layout));
		}
	}
	wofstream write(outputFile);
	write << content->c_str();
	write.close();
}
				
				
	

