import * as React from "react";
import { Component } from "react";
import styles from "./DocumentCrud.module.scss";

export interface DivingProps {
  id: number;
  name: string;
  documentLink: string;
}

export interface DivingState {}

export class Diving extends React.Component<DivingProps, DivingState> {
  public state = { title: this.props.children };

  //   public divStyle = {
  //     backgroundColor: "#f3f2f1",
  //     border: "1px solid #c8c6c4",
  //     padding: 15
  //   };
  public render() {
    // console.log(this.props.title);

    return (
      <div className="ms-Grid" dir="ltr">
        <div className="ms-Grid-row">
          <div
            className={`ms-Grid-col ms-sm6 ms-md4 ms-lg2 ${styles.divStyle}`}
          >
            {this.props.id}
          </div>
          <div
            className={`ms-Grid-col ms-sm6 ms-md8 ms-lg10 ${styles.divStyle}`}
          >
            <a
              style={{ textDecoration: "none", color: "black" }}
              href={this.props.documentLink}
              target="_blank"
            >
              {this.props.name}
            </a>
            {"  "}
            {this.props.children}
          </div>
        </div>
      </div>
    );
  }
}

// export default Diving;
