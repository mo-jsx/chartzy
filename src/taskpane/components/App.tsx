import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global console, Excel, require  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  click = async () => {
    try {
      await Excel.run(async (context) => {
        //Set env variables
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        const populationTable = currentWorksheet.tables.getItem("PopulationTable");
        const dataRange = populationTable.getDataBodyRange();

        //Create chart
        const chart = currentWorksheet.charts.add("ColumnClustered", dataRange, "Auto");

        //Set the styles and postion of the chart
        chart.title.text = "World Population";
        chart.legend.position = "Top";
        chart.setPosition("A10", "G20");
        chart.legend.format.fill.setSolidColor("blue");
        chart.dataLabels.format.font.size = 16;
        chart.dataLabels.format.font.color = "black";
        chart.series.getItemAt(0).name = "Value in Md";
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Select the source files, then click <b>Run</b>.
          </p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Create Chart
          </DefaultButton>
        </HeroList>
      </div>
    );
  }
}
