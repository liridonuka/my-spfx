import * as React from "react";
import styles from "./SearchMask.module.scss";
import { ISearchMaskProps } from "./ISearchMaskProps";
import { escape, times, isEqual } from "@microsoft/sp-lodash-subset";
import { ISearchMaskState } from "./ISearchMaskState";
import { Web, ItemAddResult, sp } from "@pnp/sp";
import { Dropdown } from "office-ui-fabric-react/lib/Dropdown";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { Rating, RatingSize } from "office-ui-fabric-react/lib/Rating";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import {
  PrimaryButton,
  DefaultButton
} from "office-ui-fabric-react/lib/Button";
import {
  Dialog,
  DialogType,
  DialogFooter
} from "office-ui-fabric-react/lib/Dialog";
import { IListItem, IPolicyUser } from "./IListItem";
import { arraysEqual } from "@uifabric/utilities/lib";
export default class DocumentCrud extends React.Component<
  ISearchMaskProps,
  ISearchMaskState
> {
  // public state = {
  //   status: this.listNotConfigured(this.props)
  //     ? "Please configure list in Web Part properties"
  //     : "Ready",
  //   documentFiles: [],
  //   policyCategories: [],
  //   policyCategoryDropDown: []
  // };
  constructor(props: ISearchMaskProps, state: ISearchMaskState) {
    super(props);

    this.state = {
      status: this.listNotConfigured(this.props)
        ? "Please configure list in Web Part properties"
        : "Ready",
      statusIndicator: 1,
      internalPolicies: [],
      documentFiles: [],
      policyUser: [],
      joinPolicyCategoryItems: [],
      policyCategoryDropDown: [],
      stringPolicyCategory: [],
      joinRegulatoryTopicItems: [],
      regulatoryTopicDropDown: [],
      stringRegulatoryTopic: [],
      joinYearItems: [],
      yearDropDown: [],
      stringYear: [],
      joinMonthItems: [],
      monthDropDown: [],
      stringMonth: [],
      anyPolicyCategorySelected: false,
      anyRegulatoryTopicSelected: false,
      anyYearSelected: false,
      anyMonthSelected: false,
      hideDialog: true,
      commentState: "",
      policyNumber: 0,
      showPanel: false
    };
  }
  public componentWillReceiveProps(nextProps: ISearchMaskProps): void {
    this.setState({
      status: this.listNotConfigured(nextProps)
        ? "Please configure list in Web Part properties"
        : "Ready",
      documentFiles: []
    });
  }

  public render(): React.ReactElement<ISearchMaskProps> {
    return (
      <div className={styles.searchMask}>
        <div>
          <div>
            <div className={styles.droping}>
              <div className={styles.dropOne}>
                <Dropdown
                  placeHolder="Filter by business function"
                  onChanged={this.filterByPolicyCategory.bind(this)}
                  multiSelect
                  options={this.state.policyCategoryDropDown}
                  // title={this.state.titleCategory}
                />
              </div>
              <div className={styles.dropTwo}>
                <Dropdown
                  placeHolder="Filter by regulatory topic"
                  onChanged={this.filterByRegulatoryTopic.bind(this)}
                  multiSelect
                  options={this.state.regulatoryTopicDropDown}
                  // title={this.state.titleCategory}
                />
              </div>
            </div>
            <div className={styles.droping}>
              <div className={styles.dropOne}>
                <Dropdown
                  placeHolder="Filter by approval year"
                  onChanged={this.filterByYear.bind(this)}
                  multiSelect
                  options={this.state.yearDropDown}
                  // title={this.state.titleCategory}
                />
              </div>
              <div className={styles.dropTwo}>
                <Dropdown
                  placeHolder="Filter by approval month"
                  onChanged={this.filterByMonth.bind(this)}
                  multiSelect
                  options={this.state.monthDropDown}
                  // title={this.state.titleCategory}
                />
              </div>
            </div>
          </div>
          {this.state.statusIndicator === 0 ? (
            <div
              style={{
                position: "fixed",
                paddingLeft: "370px",
                paddingTop: "150px",
                color: "#393939"
              }}
            >
              <p>Wait...</p>
              <Spinner size={SpinnerSize.large} />
            </div>
          ) : (
            undefined
          )}
          <div
            className={
              this.state.statusIndicator === 0 ? styles.divDisabled : undefined
            }
          >
            <div className={styles.statusDivStyle}>
              {this.state.status}{" "}
              <a href="#" onClick={() => this._showPanel()}>
                {" "}
                Show
              </a>
            </div>

            {this.state.documentFiles.slice(0, 3).map(document => (
              <div className={styles.rowDivStyle}>
                <div className={styles.policyInline}>
                  <Icon
                    onClick={
                      document.Favorite === 1 && document.Rate === 0
                        ? this.removeFromFavorites.bind(
                            this,
                            document.Name,
                            document.Id,
                            this.state.documentFiles,
                            document.Favorite === 1 ? 0 : 1
                          )
                        : document.Favorite === 1 && document.Rate > 0
                        ? this.updateFavorites.bind(
                            this,
                            document.Name,
                            document.Id,
                            this.state.documentFiles,
                            0
                          )
                        : document.Favorite === 0 && document.Rate > 0
                        ? this.updateFavorites.bind(
                            this,
                            document.Name,
                            document.Id,
                            this.state.documentFiles,
                            1
                          )
                        : this.addToFavorite.bind(
                            this,
                            document.Name,
                            document.Id,
                            document.DocumentLink,
                            this.state.documentFiles
                          )
                    }
                    iconName={
                      document.Favorite === 1
                        ? "SingleBookmarkSolid"
                        : "AddBookmark"
                    }
                    title="Add to bookmark"
                    style={{ cursor: "pointer" }}
                  />
                  &nbsp;
                  <Icon
                    onClick={() => this.entryView(document.Id)}
                    iconName="EntryView"
                    title="Policy details"
                    style={{ cursor: "pointer" }}
                  />
                </div>
                &nbsp;
                <div className={`${styles.policyInline} ${styles.divGlimmer}`}>
                  <Icon
                    title="New policy"
                    iconName={
                      parseFloat(document.Version) <= 1
                        ? document.NewDocumentExpired < 7
                          ? "glimmer"
                          : undefined
                        : undefined
                    }
                    className={styles.iconGlimmer}
                  />
                </div>
                <div className={styles.policyInline}>
                  <a
                    className={styles.linkPolicies}
                    href={document.DocumentLink}
                  >
                    {document.PolicyNumber} {document.Name} {" v"}
                    {document.Version}
                  </a>
                </div>
                <div>
                  <div title="Approved date" className={styles.policyInline}>
                    <Icon iconName="Calendar" className={styles.iconCalendar} />
                    &nbsp;
                    <div
                      className={`${styles.policyInline} ${styles.textApprovedDate}`}
                    >
                      {document.ApprovedDate}
                    </div>
                  </div>
                  &nbsp;
                  <div className={`${styles.policyInline} ${styles.rateDiv}`}>
                    <Rating
                      min={0}
                      max={5}
                      rating={document.Rate}
                      onChanged={
                        document.Rate || document.Favorite
                          ? this.updateRate.bind(
                              this,
                              document.Name,
                              document.Id,
                              document.DocumentLink,
                              this.state.documentFiles,
                              document.Comment
                            )
                          : this.addRating.bind(
                              this,
                              document.Name,
                              document.Id,
                              document.DocumentLink,
                              this.state.documentFiles
                            )
                      }
                    />
                  </div>
                  <div>
                    {/* <DefaultButton
                      secondaryText="Opens the Sample Dialog"
                      onClick={() => this._showDialog()}
                      text="Open Dialog"
                    /> */}
                    <Dialog
                      hidden={this.state.hideDialog}
                      onDismiss={() => this._closeDialog()}
                      dialogContentProps={{
                        type: DialogType.largeHeader,
                        title: "Policy comment",
                        subText:
                          "Your Inbox has changed. No longer does it include favorites, it is a singular destination for your emails."
                      }}
                      modalProps={{
                        isBlocking: false
                      }}
                    >
                      <TextField
                        label="Standard"
                        multiline
                        rows={8}
                        onChanged={this.commentState.bind(this)}
                      />
                      <DialogFooter>
                        <DefaultButton
                          onClick={() =>
                            this.updateComent(
                              this.state.policyNumber,
                              this.state.documentFiles,
                              this.state.commentState
                            )
                          }
                          text="Comment"
                        />
                        <PrimaryButton
                          onClick={() => this._closeDialog()}
                          text="Close"
                        />
                      </DialogFooter>
                    </Dialog>
                  </div>
                </div>
              </div>
            ))}
            <Panel
              isOpen={this.state.showPanel}
              type={PanelType.smallFluid}
              // onDismiss={() => this._hidePanel()}
              headerText="Panel - Small, right-aligned, fixed"
            >
              <div className={styles.searchMask}>
                {this.state.documentFiles.map(document => (
                  <div className={styles.rowDivStyle}>
                    <div className={styles.policyInline}>
                      <Icon
                        onClick={
                          document.Favorite === 1 && document.Rate === 0
                            ? this.removeFromFavorites.bind(
                                this,
                                document.Name,
                                document.Id,
                                this.state.documentFiles,
                                document.Favorite === 1 ? 0 : 1
                              )
                            : document.Favorite === 1 && document.Rate > 0
                            ? this.updateFavorites.bind(
                                this,
                                document.Name,
                                document.Id,
                                this.state.documentFiles,
                                0
                              )
                            : document.Favorite === 0 && document.Rate > 0
                            ? this.updateFavorites.bind(
                                this,
                                document.Name,
                                document.Id,
                                this.state.documentFiles,
                                1
                              )
                            : this.addToFavorite.bind(
                                this,
                                document.Name,
                                document.Id,
                                document.DocumentLink,
                                this.state.documentFiles
                              )
                        }
                        iconName={
                          document.Favorite === 1
                            ? "SingleBookmarkSolid"
                            : "AddBookmark"
                        }
                        title="Add to bookmark"
                        style={{ cursor: "pointer" }}
                      />
                      &nbsp;
                      <Icon
                        onClick={() => this.entryView(document.Id)}
                        iconName="EntryView"
                        title="Policy details"
                        style={{ cursor: "pointer" }}
                      />
                    </div>
                    &nbsp;
                    <div
                      className={`${styles.policyInline} ${styles.divGlimmer}`}
                    >
                      <Icon
                        title="New policy"
                        iconName={
                          parseFloat(document.Version) <= 1
                            ? document.NewDocumentExpired < 7
                              ? "glimmer"
                              : undefined
                            : undefined
                        }
                        className={styles.iconGlimmer}
                      />
                    </div>
                    <div className={styles.policyInline}>
                      <a
                        className={styles.linkPolicies}
                        href={document.DocumentLink}
                      >
                        {document.PolicyNumber} {document.Name} {" v"}
                        {document.Version}
                      </a>
                    </div>
                    <div>
                      <div
                        title="Approved date"
                        className={styles.policyInline}
                      >
                        <Icon
                          iconName="Calendar"
                          className={styles.iconCalendar}
                        />
                        &nbsp;
                        <div
                          className={`${styles.policyInline} ${styles.textApprovedDate}`}
                        >
                          {document.ApprovedDate}
                        </div>
                      </div>
                      &nbsp;
                      <div
                        className={`${styles.policyInline} ${styles.rateDiv}`}
                      >
                        <Rating
                          min={0}
                          max={5}
                          rating={document.Rate}
                          onChanged={
                            document.Rate || document.Favorite
                              ? this.updateRate.bind(
                                  this,
                                  document.Name,
                                  document.Id,
                                  document.DocumentLink,
                                  this.state.documentFiles,
                                  document.Comment
                                )
                              : this.addRating.bind(
                                  this,
                                  document.Name,
                                  document.Id,
                                  document.DocumentLink,
                                  this.state.documentFiles
                                )
                          }
                        />
                      </div>
                      <div>
                        {/* <DefaultButton
                      secondaryText="Opens the Sample Dialog"
                      onClick={() => this._showDialog()}
                      text="Open Dialog"
                    /> */}
                        <Dialog
                          hidden={this.state.hideDialog}
                          onDismiss={() => this._closeDialog()}
                          dialogContentProps={{
                            type: DialogType.largeHeader,
                            title: "Policy comment",
                            subText:
                              "Your Inbox has changed. No longer does it include favorites, it is a singular destination for your emails."
                          }}
                          modalProps={{
                            isBlocking: false
                          }}
                        >
                          <TextField
                            label="Standard"
                            multiline
                            rows={8}
                            onChanged={this.commentState.bind(this)}
                          />
                          <DialogFooter>
                            <DefaultButton
                              onClick={() =>
                                this.updateComent(
                                  this.state.policyNumber,
                                  this.state.documentFiles,
                                  this.state.commentState
                                )
                              }
                              text="Comment"
                            />
                            <PrimaryButton
                              onClick={() => this._closeDialog()}
                              text="Close"
                            />
                          </DialogFooter>
                        </Dialog>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </Panel>
          </div>
        </div>
      </div>
    );
  }
  private commentState(commentState): void {
    this.setState({ commentState });
  }
  public componentWillMount() {
    this.connectAndReadPolicies();
    this.connectAndReadPolicyUser();
    // this.getCurretUser();
  }
  private updateComent(policyNumber, policies, comment): void {
    this.setState({ statusIndicator: 0 });
    const selectedPolicy = this.connectAndReadPolicyUserById(policyNumber);
    selectedPolicy.then(selected => {
      let web = new Web(this.props.context.pageContext.web.absoluteUrl);
      web.lists
        .getByTitle("PolicyUser")
        .items.getById(selected)
        .update({
          Comment: comment
        })
        .then(
          (result: ItemAddResult): void => {
            // const item: IPolicyUser = result.data as IPolicyUser;
            const policyUser = this.connectAndReadPolicyUser();
            policyUser.then(policiesUser => {
              const documents = this.leftOuterJoinPolicyUser(
                policies,
                policiesUser
              );
              const rateAverage = this.connectAndReadRateAverage();
              rateAverage.then(avg => {
                const documentFiles = this.setStateAvgRate(avg, documents);
                this.setState({
                  documentFiles,
                  hideDialog: true,
                  statusIndicator: 1
                });
              });
            });
          },
          (error: any): void => {
            this.setState({
              status: "Error while adding to favorites: " + error
            });
          }
        );
    });
  }
  private updateRate(
    title,
    policyNumber,
    docLink,
    policies,
    comment,
    rate
  ): void {
    this.setState({ statusIndicator: 1 });
    const selectedPolicy = this.connectAndReadPolicyUserById(policyNumber);
    selectedPolicy.then(selected => {
      let web = new Web(this.props.context.pageContext.web.absoluteUrl);
      web.lists
        .getByTitle("PolicyUser")
        .items.getById(selected)
        .update({
          Rate: rate
        })
        .then(
          (result: ItemAddResult): void => {
            // const item: IPolicyUser = result.data as IPolicyUser;
            const policyUser = this.connectAndReadPolicyUser();
            policyUser.then(policiesUser => {
              const documents = this.leftOuterJoinPolicyUser(
                policies,
                policiesUser
              );
              const rateAverage = this.connectAndReadRateAverage();
              rateAverage.then(avg => {
                const documentFiles = this.setStateAvgRate(avg, documents);
                this.setState({
                  documentFiles,
                  statusIndicator: 1
                });
              });
            });
          },
          (error: any): void => {
            this.setState({
              status: "Error while adding to favorites: " + error
            });
          }
        );
    });
  }
  private addRating(title, policyNumber, docLink, policies, rate): void {
    this.setState({ statusIndicator: 1 });
    let web = new Web(this.props.context.pageContext.web.absoluteUrl);
    web.lists
      .getByTitle("PolicyUser")
      .items.add({
        Title: title,
        PolicyNumber: policyNumber,
        Rate: rate,
        Policy: title,
        PolicyLink: docLink
      })
      .then(
        (result: ItemAddResult): void => {
          // const item: IPolicyUser = result.data as IPolicyUser;
          const policyUser = this.connectAndReadPolicyUser();
          policyUser.then(policiesUser => {
            const documents = this.leftOuterJoinPolicyUser(
              policies,
              policiesUser
            );
            const rateAverage = this.connectAndReadRateAverage();
            rateAverage.then(avg => {
              const documentFiles = this.setStateAvgRate(avg, documents);
              this.setState({
                documentFiles,
                status: `${title} was added to favorites`,
                statusIndicator: 1
              });
            });
          });
        },
        (error: any): void => {
          this.setState({
            status: "Error while adding to favorites: " + error
          });
        }
      );
    this._showDialog(policyNumber);
  }
  private addToFavorite(title, policyNumber, docLink, policies): void {
    this.setState({ status: "Adding to favorites...", statusIndicator: 0 });
    let web = new Web(this.props.context.pageContext.web.absoluteUrl);
    web.lists
      .getByTitle("PolicyUser")
      .items.add({
        Title: title,
        PolicyNumber: policyNumber,
        Favorite: 1,
        Policy: title,
        PolicyLink: docLink
      })
      .then(
        (result: ItemAddResult): void => {
          // const item: IPolicyUser = result.data as IPolicyUser;
          const policyUser = this.connectAndReadPolicyUser();
          policyUser.then(policiesUser => {
            const documents = this.leftOuterJoinPolicyUser(
              policies,
              policiesUser
            );
            const rateAverage = this.connectAndReadRateAverage();
            rateAverage.then(avg => {
              const documentFiles = this.setStateAvgRate(avg, documents);
              this.setState({
                documentFiles,
                status: `${title} was added to favorites`,
                statusIndicator: 1
              });
            });
          });
        },
        (error: any): void => {
          this.setState({
            status: "Error while adding to favorites: " + error
          });
        }
      );
  }
  private updateFavorites(title, policyNumber, policies, favorite): void {
    this.setState({
      status:
        favorite === 1
          ? "Adding to favorites..."
          : "Removing from favorites...",
      statusIndicator: 1
    });
    const selectedPolicy = this.connectAndReadPolicyUserById(policyNumber);
    selectedPolicy.then(selected => {
      let web = new Web(this.props.context.pageContext.web.absoluteUrl);
      web.lists
        .getByTitle("PolicyUser")
        .items.getById(selected)
        .update({
          Favorite: favorite
        })
        .then(
          (result: ItemAddResult): void => {
            // const item: IPolicyUser = result.data as IPolicyUser;
            const policyUser = this.connectAndReadPolicyUser();
            policyUser.then(policiesUser => {
              const documents = this.leftOuterJoinPolicyUser(
                policies,
                policiesUser
              );
              const rateAverage = this.connectAndReadRateAverage();
              rateAverage.then(avg => {
                const documentFiles = this.setStateAvgRate(avg, documents);
                this.setState({
                  documentFiles,
                  status:
                    favorite === 1
                      ? `${title} was added to favorites`
                      : `${title} was removed from favorites`,
                  statusIndicator: 1
                });
              });
            });
          },
          (error: any): void => {
            this.setState({
              status: "Error while adding to favorites: " + error
            });
          }
        );
    });
  }
  private removeFromFavorites(title, policyNumber, policies): void {
    this.setState({ status: "Updating favorites...", statusIndicator: 0 });
    const selectedPolicy = this.connectAndReadPolicyUserById(policyNumber);
    selectedPolicy.then(selected => {
      let web = new Web(this.props.context.pageContext.web.absoluteUrl);
      web.lists
        .getByTitle("PolicyUser")
        .items.getById(selected)
        .delete()
        .then(
          result => {
            // const item: IPolicyUser = result.data as IPolicyUser;
            const policyUser = this.connectAndReadPolicyUser();
            policyUser.then(policiesUser => {
              const documents = this.leftOuterJoinPolicyUser(
                policies,
                policiesUser
              );
              const rateAverage = this.connectAndReadRateAverage();
              rateAverage.then(avg => {
                const documentFiles = this.setStateAvgRate(avg, documents);
                this.setState({
                  documentFiles,
                  status: `${title} was removed from favorites`,
                  statusIndicator: 1
                });
              });
            });
          },
          (error: any): void => {
            this.setState({
              status: "Error while adding to favorites: " + error
            });
          }
        );
    });
  }
  private connectAndReadPolicyUserById(policyNumber) {
    const ama = this.getCurretUser();
    return ama.then(userId => {
      const web = new Web(this.props.context.pageContext.web.absoluteUrl);
      return new Promise<number>(
        (
          resolve: (itemId: number) => void,
          reject: (error: any) => void
        ): void => {
          web.lists
            .getByTitle("PolicyUser")
            .items.filter(
              `AuthorId eq '${userId}' and PolicyNumber eq '${policyNumber}'`
            )
            .select("Id")
            .get()
            .then(
              (items: { Id: number }[]): void => {
                if (items.length === 0) {
                  resolve(-1);
                } else {
                  resolve(items[0].Id);
                }
              },
              (error: any): void => {
                reject(error);
              }
            );
        }
      );
    });
  }
  private connectAndReadRateAverage() {
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    return web.lists
      .getByTitle("PolicyUser")
      .items.select("PolicyNumber", "Rate")
      .get()
      .then(
        policyUser => {
          return policyUser;
        },
        (error: any): void => {
          this.setState({
            status: "Loading all items failed with error: " + error
          });
        }
      );
  }
  private connectAndReadPolicyUser() {
    const ama = this.getCurretUser();
    return ama.then(userId => {
      const web = new Web(this.props.context.pageContext.web.absoluteUrl);
      return web.lists
        .getByTitle("PolicyUser")
        .items.filter(`AuthorId eq '${userId}'`)
        .get()
        .then(
          policyUser => {
            return policyUser;
          },
          (error: any): void => {
            this.setState({
              status: "Loading all items failed with error: " + error
            });
          }
        );
    });
  }

  private connectAndReadPolicies(): void {
    this.setState({ documentFiles: [], status: "Loading all items..." });
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    web.lists
      .getByTitle(this.props.listName)
      .items.expand("File")
      .getAll()
      .then(
        items => {
          this.getPolicies(items);
        },
        (error: any): void => {
          this.setState({
            status: "Loading all items failed with error: " + error
          });
        }
      );
  }
  private getPolicies(items): void {
    const monthNames = this.monthName();
    let policies = [];
    let joinPolicyCategoryItems = [];
    let joinRegulatoryTopicItems = [];
    let itemList = [];
    items.forEach(policy => {
      itemList.push({
        Id: policy.Id,
        PolicyNumber: policy.Policy_x0020_Number,
        Name: policy.File.Name,
        Version: policy.OData__UIVersionString,
        DocumentLink: policy.ServerRedirectedEmbedUrl,
        ApprovedDate: policy.Date_x0020_of_x0020_approval
      });
      policy.Policy_x0020_Category.forEach(policyCategory => {
        joinPolicyCategoryItems.push({
          Id: policy.Id,
          PolicyCategory: policyCategory.Label.split(/:/)[1]
        });
      });
      policy.Regulatory_x0020_Topic.forEach(regulatoryTopic => {
        joinRegulatoryTopicItems.push({
          Id: policy.Id,
          RegulatoryTopic: regulatoryTopic.Label.split(/:/)[1]
        });
      });
    });
    joinPolicyCategoryItems.forEach(j => {
      itemList
        .filter(f => f.Id === j.Id)
        .map(item => (item.PolicyCategory += ";" + j.PolicyCategory));
    });
    joinRegulatoryTopicItems.forEach(j => {
      itemList
        .filter(f => f.Id === j.Id)
        .map(item => (item.RegulatoryTopic += ";" + j.RegulatoryTopic));
    });
    itemList.forEach(policy => {
      const diffDays = this.getDiffDays(policy.ApprovedDate);
      policies.push({
        Id: policy.Id,
        PolicyNumber: policy.PolicyNumber,
        Name: policy.Name,
        Version: policy.Version,
        DocumentLink: policy.DocumentLink,
        ApprovedDate: new Date(policy.ApprovedDate).toLocaleDateString(),
        PolicyCategory: policy.PolicyCategory
          ? policy.PolicyCategory.split("undefined;").pop()
          : "",
        RegulatoryTopic: policy.RegulatoryTopic
          ? policy.RegulatoryTopic.split("undefined;").pop()
          : "",
        Year: new Date(policy.ApprovedDate).getFullYear().toString(),
        Month: monthNames[new Date(policy.ApprovedDate).getMonth()],
        NewDocumentExpired: policy.ApprovedDate ? diffDays : undefined
      });
    });
    const policyUser = this.connectAndReadPolicyUser();
    policyUser.then(policiesUser => {
      const documents = this.leftOuterJoinPolicyUser(policies, policiesUser);
      const rateAverage = this.connectAndReadRateAverage();
      rateAverage.then(avg => {
        const documentFiles = this.setStateAvgRate(avg, documents);
        this.setState({
          documentFiles,
          internalPolicies: documentFiles,
          joinPolicyCategoryItems,
          joinRegulatoryTopicItems,
          status: `Successfully loaded ${documentFiles.length} items`
        });
        this.dropDownPolicyCategory(documentFiles);
        this.dropDownRegulatoryTopic(documentFiles);
        this.dropDownDateYear(documentFiles);
        this.dropDownDateMonth(documentFiles);
      });
    });
  }
  //dropdowns
  private dropDownPolicyCategory(items?): void {
    const filedName = "PolicyCategory";
    const policyCategoryDropDown = this.fillDropDown(items, filedName);
    this.setState({ policyCategoryDropDown });
  }

  private dropDownRegulatoryTopic(items?): void {
    const filedName = "RegulatoryTopic";
    const regulatoryTopicDropDown = this.fillDropDown(items, filedName);
    this.setState({ regulatoryTopicDropDown });
  }
  private dropDownDateYear(items?): void {
    const filedName = "Year";
    const yearDropDown = this.fillDropDown(items, filedName);
    this.setState({ yearDropDown });
  }
  private dropDownDateMonth(items?): void {
    const filedName = "Month";
    const monthDropDown = this.fillDropDown(items, filedName);
    this.setState({ monthDropDown });
  }
  private fillDropDown(items?, filedName?) {
    let documentFiles = [];
    items
      ? documentFiles.push(...items)
      : documentFiles.push(...this.state.documentFiles);
    let listBeforeSplit = [];
    let listNoUnique = [];
    documentFiles.forEach(item => {
      if (item[filedName]) {
        listBeforeSplit.push({
          text: ";" + item[filedName]
        });
      }
    });

    listBeforeSplit.forEach(element => {
      element.text.split(";").forEach(split => {
        if (split) {
          listNoUnique.push(split);
        }
      });
    });
    let uniqueItems = new Set(listNoUnique.map(unique => unique));
    let dropDowResult = [];
    uniqueItems.forEach(uniqueItem => {
      dropDowResult.push({ key: uniqueItem, text: uniqueItem });
    });
    return dropDowResult;
  }
  private filterByPolicyCategory(selectedItems) {
    const stringPolicyCategory = this.selectItems(
      selectedItems,
      this.state.stringPolicyCategory
    );
    this.setState({
      stringPolicyCategory,
      statusIndicator: 1,
      status: "Applying filters..."
    });
    const cloned = this.clonedList(
      this.state.anyRegulatoryTopicSelected ||
        this.state.anyYearSelected ||
        this.state.anyMonthSelected
        ? stringPolicyCategory.length > 0
          ? this.state.joinPolicyCategoryItems
          : this.state.joinPolicyCategoryItems
        : this.state.internalPolicies
    );
    const policyUser = this.connectAndReadPolicyUser();
    policyUser.then(policiesUser => {
      const clonedListWithFavorite = this.leftOuterJoinPolicyUser(
        cloned,
        policiesUser
      );
      const rateAverage = this.connectAndReadRateAverage();
      rateAverage.then(avg => {
        const clonedList = this.setStateAvgRate(avg, clonedListWithFavorite);
        let filteredList = [];
        stringPolicyCategory.forEach(s => {
          clonedList
            .filter(f => f.PolicyCategory.includes(s))
            .map(item => filteredList.push(item));
        });
        let uniqueItems = new Set(filteredList.map(unique => unique));
        let documentFiles = [];
        uniqueItems.forEach(u => {
          documentFiles.push(u);
        });
        this.setState({
          documentFiles:
            stringPolicyCategory.length > 0 ? documentFiles : clonedList,
          joinRegulatoryTopicItems:
            stringPolicyCategory.length > 0 ? documentFiles : clonedList,
          joinYearItems:
            stringPolicyCategory.length > 0 ? documentFiles : clonedList,
          joinMonthItems:
            stringPolicyCategory.length > 0 ? documentFiles : clonedList,
          anyPolicyCategorySelected:
            stringPolicyCategory.length > 0 ? true : false,
          status: `Successfully loaded ${
            stringPolicyCategory.length > 0
              ? documentFiles.length
              : clonedList.length
          } items`,
          statusIndicator: 1
        });
        if (!this.state.anyRegulatoryTopicSelected) {
          this.dropDownRegulatoryTopic(
            stringPolicyCategory.length > 0 ? documentFiles : clonedList
          );
        }
        if (!this.state.anyYearSelected) {
          this.dropDownDateYear(
            stringPolicyCategory.length > 0 ? documentFiles : clonedList
          );
        }
        if (!this.state.anyMonthSelected) {
          this.dropDownDateMonth(
            stringPolicyCategory.length > 0 ? documentFiles : clonedList
          );
        }
      });
    });
  }
  private filterByRegulatoryTopic(selectedItems) {
    const stringRegulatoryTopic = this.selectItems(
      selectedItems,
      this.state.stringRegulatoryTopic
    );
    this.setState({
      stringRegulatoryTopic,
      statusIndicator: 1,
      status: "Applying filters..."
    });
    const cloned = this.clonedList(
      this.state.anyPolicyCategorySelected ||
        this.state.anyYearSelected ||
        this.state.anyMonthSelected
        ? stringRegulatoryTopic.length > 0
          ? this.state.joinRegulatoryTopicItems
          : this.state.joinRegulatoryTopicItems
        : this.state.internalPolicies
    );
    const policyUser = this.connectAndReadPolicyUser();
    policyUser.then(policiesUser => {
      const clonedListWithFavorite = this.leftOuterJoinPolicyUser(
        cloned,
        policiesUser
      );
      const rateAverage = this.connectAndReadRateAverage();
      rateAverage.then(avg => {
        const clonedList = this.setStateAvgRate(avg, clonedListWithFavorite);
        let filteredList = [];
        stringRegulatoryTopic.forEach(s => {
          clonedList
            .filter(f => f.RegulatoryTopic.includes(s))
            .map(item => filteredList.push(item));
        });
        let uniqueItems = new Set(filteredList.map(unique => unique));
        let documentFiles = [];
        uniqueItems.forEach(u => {
          documentFiles.push(u);
        });
        this.setState({
          documentFiles:
            stringRegulatoryTopic.length > 0 ? documentFiles : clonedList,
          joinPolicyCategoryItems:
            stringRegulatoryTopic.length > 0 ? documentFiles : clonedList,
          joinYearItems:
            stringRegulatoryTopic.length > 0 ? documentFiles : clonedList,
          joinMonthItems:
            stringRegulatoryTopic.length > 0 ? documentFiles : clonedList,
          anyRegulatoryTopicSelected:
            stringRegulatoryTopic.length > 0 ? true : false,
          status: `Successfully loaded ${
            stringRegulatoryTopic.length > 0
              ? documentFiles.length
              : clonedList.length
          } items`,
          statusIndicator: 1
        });

        if (!this.state.anyPolicyCategorySelected) {
          this.dropDownPolicyCategory(
            stringRegulatoryTopic.length > 0 ? documentFiles : clonedList
          );
        }
        if (!this.state.anyYearSelected) {
          this.dropDownDateYear(
            stringRegulatoryTopic.length > 0 ? documentFiles : clonedList
          );
        }
        if (!this.state.anyMonthSelected) {
          this.dropDownDateMonth(
            stringRegulatoryTopic.length > 0 ? documentFiles : clonedList
          );
        }
      });
    });
  }
  private filterByYear(selectedItems) {
    const stringYear = this.selectItems(selectedItems, this.state.stringYear);
    this.setState({
      stringYear,
      statusIndicator: 1,
      status: "Applying filters..."
    });
    const cloned = this.clonedList(
      this.state.anyPolicyCategorySelected ||
        this.state.anyRegulatoryTopicSelected ||
        this.state.anyMonthSelected
        ? stringYear.length > 0
          ? this.state.joinYearItems
          : this.state.joinYearItems
        : this.state.internalPolicies
    );
    const policyUser = this.connectAndReadPolicyUser();
    policyUser.then(policiesUser => {
      const clonedListWithFavorite = this.leftOuterJoinPolicyUser(
        cloned,
        policiesUser
      );
      const rateAverage = this.connectAndReadRateAverage();
      rateAverage.then(avg => {
        const clonedList = this.setStateAvgRate(avg, clonedListWithFavorite);
        let filteredList = [];
        stringYear.forEach(s => {
          clonedList
            .filter(f => f.Year.includes(s))
            .map(item => filteredList.push(item));
        });
        let uniqueItems = new Set(filteredList.map(unique => unique));
        let documentFiles = [];
        uniqueItems.forEach(u => {
          documentFiles.push(u);
        });
        this.setState({
          documentFiles: stringYear.length > 0 ? documentFiles : clonedList,
          joinPolicyCategoryItems:
            stringYear.length > 0 ? documentFiles : clonedList,
          joinRegulatoryTopicItems:
            stringYear.length > 0 ? documentFiles : clonedList,
          joinMonthItems: stringYear.length > 0 ? documentFiles : clonedList,
          anyYearSelected: stringYear.length > 0 ? true : false,
          status: `Successfully loaded ${
            stringYear.length > 0 ? documentFiles.length : clonedList.length
          } items`,
          statusIndicator: 1
        });

        if (!this.state.anyPolicyCategorySelected) {
          this.dropDownPolicyCategory(
            stringYear.length > 0 ? documentFiles : clonedList
          );
        }
        if (!this.state.anyRegulatoryTopicSelected) {
          this.dropDownRegulatoryTopic(
            stringYear.length > 0 ? documentFiles : clonedList
          );
        }
        if (!this.state.anyMonthSelected) {
          this.dropDownDateMonth(
            stringYear.length > 0 ? documentFiles : clonedList
          );
        }
      });
    });
  }
  private filterByMonth(selectedItems) {
    const stringMonth = this.selectItems(selectedItems, this.state.stringMonth);
    this.setState({
      stringMonth,
      statusIndicator: 1,
      status: "Applying filters..."
    });
    const cloned = this.clonedList(
      this.state.anyPolicyCategorySelected ||
        this.state.anyRegulatoryTopicSelected ||
        this.state.anyYearSelected
        ? stringMonth.length > 0
          ? this.state.joinMonthItems
          : this.state.joinMonthItems
        : this.state.internalPolicies
    );
    const policyUser = this.connectAndReadPolicyUser();
    policyUser.then(policiesUser => {
      const clonedListWithFavorite = this.leftOuterJoinPolicyUser(
        cloned,
        policiesUser
      );
      const rateAverage = this.connectAndReadRateAverage();
      rateAverage.then(avg => {
        const clonedList = this.setStateAvgRate(avg, clonedListWithFavorite);
        let filteredList = [];
        stringMonth.forEach(s => {
          clonedList
            .filter(f => f.Month.includes(s))
            .map(item => filteredList.push(item));
        });
        let uniqueItems = new Set(filteredList.map(unique => unique));
        let documentFiles = [];
        uniqueItems.forEach(u => {
          documentFiles.push(u);
        });
        this.setState({
          documentFiles: stringMonth.length > 0 ? documentFiles : clonedList,
          joinPolicyCategoryItems:
            stringMonth.length > 0 ? documentFiles : clonedList,
          joinRegulatoryTopicItems:
            stringMonth.length > 0 ? documentFiles : clonedList,
          joinYearItems: stringMonth.length > 0 ? documentFiles : clonedList,
          anyMonthSelected: stringMonth.length > 0 ? true : false,
          status: `Successfully loaded ${
            stringMonth.length > 0 ? documentFiles.length : clonedList.length
          } items`,
          statusIndicator: 1
        });
        if (!this.state.anyPolicyCategorySelected) {
          this.dropDownPolicyCategory(
            stringMonth.length > 0 ? documentFiles : clonedList
          );
        }
        if (!this.state.anyRegulatoryTopicSelected) {
          this.dropDownRegulatoryTopic(
            stringMonth.length > 0 ? documentFiles : clonedList
          );
        }
        if (!this.state.anyYearSelected) {
          this.dropDownDateYear(
            stringMonth.length > 0 ? documentFiles : clonedList
          );
        }
      });
    });
  }
  private selectItems(selectedItems, stringList) {
    let result = [...stringList];
    if (selectedItems.selected) {
      // add the option if it's checked
      result.push(selectedItems.text);
    } else {
      // remove the option if it's unchecked
      const currIndex = result.indexOf(selectedItems.text);
      if (currIndex > -1) {
        result.splice(currIndex, 1);
      }
    }
    return result;
  }
  private clonedList(sourceList) {
    let clonedList = [...sourceList];
    let list = [];
    clonedList.forEach(policy => {
      list.push({
        Id: policy.Id,
        PolicyNumber: policy.PolicyNumber,
        Name: policy.Name,
        Version: policy.Version,
        DocumentLink: policy.DocumentLink,
        ApprovedDate: new Date(policy.ApprovedDate).toLocaleDateString(),
        PolicyCategory: policy.PolicyCategory,
        RegulatoryTopic: policy.RegulatoryTopic,
        Year: policy.Year,
        Month: policy.Month,
        NewDocumentExpired: policy.NewDocumentExpired,
        Favorite: policy.Favorite
      });
    });
    return list;
  }
  private monthName() {
    const monthNames = [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December"
    ];
    return monthNames;
  }
  private getDiffDays(elementDate) {
    const date1 = new Date(elementDate);
    const date2 = new Date();
    const diffTime = Math.abs(date2.getTime() - date1.getTime());
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays;
  }
  private entryView(id): void {
    const url = `${
      this.props.context.pageContext.web.absoluteUrl
    }/${this.props.listName.replace(/ +/g, "")}/Forms/DispForm.aspx?ID=${id}`;
    window.open(url, "_blank");
  }
  private getCurretUser() {
    const userEmail = this.props.context.pageContext.user.email;
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    return web.ensureUser(userEmail).then(result => {
      return result.data.Id;
    });
  }
  private leftOuterJoinPolicyUser(a, b) {
    let result = [];
    a.forEach(policy => {
      let found = false;
      b.forEach(policyUser => {
        if (policy.Id === policyUser.PolicyNumber) {
          result.push({
            Id: policy.Id,
            PolicyNumber: policy.PolicyNumber,
            Name: policy.Name,
            Version: policy.Version,
            DocumentLink: policy.DocumentLink,
            ApprovedDate: policy.ApprovedDate,
            PolicyCategory: policy.PolicyCategory,
            RegulatoryTopic: policy.RegulatoryTopic,
            Year: policy.Year,
            Month: policy.Month,
            NewDocumentExpired: policy.NewDocumentExpired,
            Favorite: policyUser.Favorite,
            Rate: policyUser.Rate
          });
          found = true;
        }
      });

      if (found === false) {
        result.push({
          Id: policy.Id,
          PolicyNumber: policy.PolicyNumber,
          Name: policy.Name,
          Version: policy.Version,
          DocumentLink: policy.DocumentLink,
          ApprovedDate: policy.ApprovedDate,
          PolicyCategory: policy.PolicyCategory,
          RegulatoryTopic: policy.RegulatoryTopic,
          Year: policy.Year,
          Month: policy.Month,
          NewDocumentExpired: policy.NewDocumentExpired,
          Favorite: 0,
          Rate: 0
        });
      }
    });
    return result;
  }
  private setStateAvgRate(avgResponse, result) {
    let avgList = [];
    avgResponse.forEach(element => {
      avgList.push({
        PolicyNumber: element.PolicyNumber,
        Rate: element.Rate
      });
    });

    let groupeData = avgList.reduce((l, r) => {
      let key = r.PolicyNumber;
      if (typeof l[key] === "undefined") {
        l[key] = {
          sum: 0,
          count: 0
        };
      }
      l[key].sum += r.Rate;
      l[key].count += 1;
      return l;
    }, {});

    let avgGroupedData = Object.keys(groupeData).map(key => {
      let keyParts = key.split(/\|/);
      return {
        PolicyNumber: parseInt(keyParts[0], 10),
        Rate: groupeData[key].sum / groupeData[key].count
      };
    });
    let results = [];
    for (let x = 0; x < result.length; x++) {
      let foundX = false;
      if (result[x].Rate) {
        for (let y = 0; y < avgGroupedData.length; y++) {
          if (result[x].Id === avgGroupedData[y].PolicyNumber) {
            results.push({
              Id: result[x].Id,
              PolicyNumber: result[x].PolicyNumber,
              Name: result[x].Name,
              Version: result[x].Version,
              DocumentLink: result[x].DocumentLink,
              ApprovedDate: result[x].ApprovedDate,
              PolicyCategory: result[x].PolicyCategory,
              RegulatoryTopic: result[x].RegulatoryTopic,
              Year: result[x].Year,
              Month: result[x].Month,
              Favorite: result[x].Favorite,
              Rate: avgGroupedData[y].Rate,
              Comment: result[x].Comment
            });
            foundX = true;
            break;
          }
        }
        if (foundX === false) {
          results.push({
            Id: result[x].Id,
            PolicyNumber: result[x].PolicyNumber,
            Name: result[x].Name,
            Version: result[x].Version,
            DocumentLink: result[x].DocumentLink,
            ApprovedDate: result[x].ApprovedDate,
            PolicyCategory: result[x].PolicyCategory,
            RegulatoryTopic: result[x].RegulatoryTopic,
            Year: result[x].Year,
            Month: result[x].Month,
            Favorite: result[x].Favorite,
            Rate: 0,
            Comment: result[x].Comment
          });
        }
      } else {
        results.push({
          Id: result[x].Id,
          PolicyNumber: result[x].PolicyNumber,
          Name: result[x].Name,
          Version: result[x].Version,
          DocumentLink: result[x].DocumentLink,
          ApprovedDate: result[x].ApprovedDate,
          PolicyCategory: result[x].PolicyCategory,
          RegulatoryTopic: result[x].RegulatoryTopic,
          Year: result[x].Year,
          Month: result[x].Month,
          Favorite: result[x].Favorite,
          Rate: 0,
          Comment: result[x].Comment
        });
      }
    }

    return results;
  }
  private _showDialog(policyNumber): void {
    this.setState({ hideDialog: false, policyNumber });
  }

  private _closeDialog(): void {
    this.setState({ hideDialog: true });
  }
  private _showPanel(): void {
    console.log("Hit");
    this.setState({ showPanel: true });
  }

  private _hidePanel(): void {
    this.setState({ showPanel: false });
  }

  private listNotConfigured(props: ISearchMaskProps): boolean {
    return (
      props.listName === undefined ||
      props.listName === null ||
      props.listName.length === 0
    );
  }
}
