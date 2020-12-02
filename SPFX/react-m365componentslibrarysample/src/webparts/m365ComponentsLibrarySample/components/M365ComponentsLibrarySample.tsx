import * as React from 'react';
import styles from './M365ComponentsLibrarySample.module.scss';
import { AdvancedDetailsList, CurrencyTextField, CurrencyType, IView, M365ComponentsLibrary } from 'react-m-365-componentslibrary';
import { DefaultButton, PrimaryButton, Stack, StackItem, TextField } from 'office-ui-fabric-react';

let detailsListItems: any[] = [
  {
    id: 1,
    title: "item 1",
    created: new Date(2020, 6, 15, 12, 15),
    location: "Nantes",
    author: {
      name: "Robert"
    },
    statut: "Validé"
  },
  {
    id: 2,
    title: "item 2",
    created: new Date(2020, 5, 3, 14, 37),
    location: "Nantes",
    author: {
      name: "Gérard"
    },
    statut: "En attente"
  },
  {
    id: 3,
    title: "item 3",
    created: new Date(2020, 9, 4, 16, 52),
    location: "Nantes",
    author: {
      name: "Robert"
    },
    statut: "Brouillon"
  },
  {
    id: 4,
    title: "item 4",
    created: new Date(2020, 6, 15, 12, 15),
    location: "Paris",
    author: {
      name: "Robert"
    },
    statut: "En attente"
  },
  {
    id: 5,
    title: "item 5",
    created: new Date(2020, 4, 2, 9, 37),
    location: "Nantes",
    author: {
      name: "Gérard"
    },
    statut: "Validé"
  },
  {
    id: 6,
    title: "item 6",
    created: new Date(2020, 9, 4, 16, 23),
    location: "Paris",
    author: {
      name: "Marie"
    },
    statut: "En attente"
  },
  {
    id: 7,
    title: "item 7",
    created: new Date(2020, 6, 15, 12, 15),
    location: "Lyon",
    author: {
      name: "Robert"
    },
    statut: "Validé"
  },
  {
    id: 8,
    title: "item 8",
    created: new Date(2020, 4, 9, 10, 4),
    location: "Lyon",
    author: {
      name: "Jeanne"
    },
    statut: "Validé"
  },
  {
    id: 9,
    title: "item 9",
    created: new Date(2020, 12, 4, 8, 21),
    location: "Paris",
    author: {
      name: "Jeanne"
    },
    statut: "Brouillon"
  }
];

const defaultView: IView = {
  fields: [
    {
      displayName: "Titre",
      fieldName: "title"
    },
    {
      displayName: "Créé le",
      fieldName: "created",
      type: "datetime"
    },
    {
      displayName: "Créé par",
      fieldName: "author.name"
    },
    {
      displayName: "Ville",
      fieldName: "location",
      isHidden: true
    },
    {
      displayName: "Statut",
      fieldName: "statut",
      onRender: (item) => {
        return (
          <div style={
            {
              backgroundColor: item.statut === "Validé" ? "#B4FA96" : (item.statut === "En attente" ? "#F7BA8D" : "inherit"),
              color: item.statut === "Validé" ? "#32D939" : (item.statut === "En attente" ? "#DF6308" : "inherit"),
              textAlign: "center"
            }
          }>
            {item.statut}
          </div>);
      }
    }
  ],
  filteredBy: [],
  sortedBy: [{ property: "author.name", isSortedDescending: true }],
  groupedBy: [{ fieldName: "author.name", isCollapsed: false }]
};

export interface IM365ComponentsLibrarySampleState {
  detailsListItems: any[];
  formPanelItem: any;
}

export default class M365ComponentsLibrarySample extends React.Component<{}, IM365ComponentsLibrarySampleState> {

  constructor(props) {
    super(props);
    this.state = {
      detailsListItems: detailsListItems,
      formPanelItem: null
    };
  }

  public render(): React.ReactElement<{}> {
    return (
      <div className={styles.m365ComponentsLibrarySample}>
        <Stack>
          <StackItem>
            <Stack horizontal horizontalAlign={"space-evenly"}>
              <CurrencyTextField
                borderless={false}
                currency={CurrencyType.Euro}
                description={"Saisir un montant inférieur à 100€."}
                label={"Montant avec validation"}
                onBlur={(value) => { console.log(value); }}
                onGetErrorMessage={
                  (text) => {
                    let value = parseFloat(text);
                    if (value >= 100) {
                      return 'Montant non valide.';
                    }
                  }
                }
                value={5} />
            </Stack>
          </StackItem> 
          <StackItem>
            <hr />
          </StackItem>
          <StackItem>
            <Stack>
              <AdvancedDetailsList
                items={this.state.detailsListItems}
                defaultView={defaultView}
                disableFilterBy={false}
                disableGroupBy={false}
                disableSortBy={false}
                disableCommandBar={false}
                onRemoveItems={null}
                onRenderFormPanel={
                  (item, closePanel) => {
                    return (
                      <Stack>
                        <StackItem>
                          <TextField
                            label={"Titre"}
                            value={this.state.formPanelItem && this.state.formPanelItem.title ? this.state.formPanelItem.title : (item && item.title ? item.title : "")}
                            onBlur={
                              (event) => {
                                const newValue = event.target.value;
                                let tmpFormPanelItem = this.state.formPanelItem;
                                if (!tmpFormPanelItem) {
                                  tmpFormPanelItem = item ?
                                    item : {
                                      title: "",
                                      created: null,
                                      author: {
                                        name: ""
                                      },
                                      statut: "Brouillon"
                                    };
                                }
                                tmpFormPanelItem.title = newValue;
                                this.setState({
                                  formPanelItem: tmpFormPanelItem
                                });
                              }
                            } />
                        </StackItem>
                        <StackItem>
                          <TextField
                            label={"Auteur"}
                            value={this.state.formPanelItem && this.state.formPanelItem.author ? this.state.formPanelItem.author.name : (item && item.author ? item.author.name : "")}
                            onBlur={
                              (event) => {
                                const newValue = event.target.value;
                                let tmpFormPanelItem = this.state.formPanelItem;
                                if (!tmpFormPanelItem) {
                                  tmpFormPanelItem = item ?
                                    item : {
                                      title: "",
                                      created: null,
                                      author: {
                                        name: ""
                                      },
                                      statut: "Brouillon"
                                    };
                                }
                                tmpFormPanelItem.author = {
                                  name: newValue
                                };
                                this.setState({
                                  formPanelItem: tmpFormPanelItem
                                });
                              }
                            } />
                        </StackItem>
                        <StackItem>
                          <br />
                        </StackItem>
                        <StackItem>
                          <Stack horizontal horizontalAlign={"space-evenly"}>
                            <DefaultButton text={"Fermer"} onClick={closePanel} />
                            <PrimaryButton
                              text={"Enregistrer"}
                              onClick={
                                () => {
                                  let tmpDetailsListItems = this.state.detailsListItems.slice();
                                  let tmpFormPanelItem = this.state.formPanelItem;
                                  if (tmpFormPanelItem.id) {
                                    for (let i = 0; i < tmpDetailsListItems.length; i++) {
                                      if (tmpFormPanelItem.id === tmpDetailsListItems[i].id) {
                                        tmpDetailsListItems[i] = tmpFormPanelItem;
                                        break;
                                      }
                                    }
                                  }
                                  else {
                                    tmpFormPanelItem.id = tmpDetailsListItems.length + 1;
                                    tmpDetailsListItems.push(tmpFormPanelItem);
                                  }
                                  this.setState(
                                    {
                                      detailsListItems: tmpDetailsListItems,
                                      formPanelItem: null
                                    }
                                  );
                                  console.log("Enregistrement!");
                                  closePanel();
                                }
                              }
                              disabled={!this.state.formPanelItem} />
                          </Stack>
                        </StackItem>
                      </Stack>
                    );
                  }
                }
              />
            </Stack>
          </StackItem>
        </Stack>
      </div>
    );
  }
}
