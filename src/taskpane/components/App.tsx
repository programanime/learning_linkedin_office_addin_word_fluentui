import * as React from "react";
import { Coachmark, DefaultButton, DirectionalHint, IButtonProps, Icon, ILabelStyles, IPivotItemProps, IStyleSet, Label, Pivot, PivotItem, TeachingBubble, TeachingBubbleContent } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { MessageBar, MessageBarType, Toggle, Text, mergeStyles } from '@fluentui/react';
import { Nav, INavLink, INavStyles, INavLinkGroup } from '@fluentui/react/lib/Nav';


const navStyles: Partial<INavStyles> = {
  root: {
    width: 208,
    height: 350,
    boxSizing: 'border-box',
    border: '1px solid #eee',
    overflowY: 'auto',
  },
};

const navLinkGroups: INavLinkGroup[] = [
  {
    links: [
      {
        name: 'Home',
        url: 'http://example.com',
        expandAriaLabel: 'Expand Home section',
        links: [
          {
            name: 'Activity',
            url: 'http://msn.com',
            key: 'key1',
            target: '_blank',
          },
          {
            name: 'MSN',
            url: 'http://msn.com',
            disabled: true,
            key: 'key2',
            target: '_blank',
          },
        ],
        isExpanded: false,
      },
      {
        name: 'Documents',
        url: 'http://example.com',
        key: 'key3',
        isExpanded: true,
        target: '_blank',
      },
      {
        name: 'Pages',
        url: 'http://msn.com',
        key: 'key4',
        target: '_blank',
      },
      {
        name: 'Notebook',
        url: 'http://msn.com',
        key: 'key5',
        disabled: true,
      },
      {
        name: 'Communication and Media',
        url: 'http://msn.com',
        key: 'key6',
        target: '_blank',
      },
      {
        name: 'News',
        url: 'http://cnn.com',
        icon: 'News',
        key: 'key7',
        target: '_blank',
      },
    ],
  },
];

const buttonProps: IButtonProps = {
  text: 'Try it',
};

const buttonProps2: IButtonProps = {
  text: 'Try it again',
};

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

const examplePrimaryButtonProps: IButtonProps = {
  children: 'Try it out',
};

const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
  root: { marginTop: 10 },
};

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
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
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

    const botoncitoCoachmark = document.getElementById("botoncitoCoachmark");

    return (
      <>
         <div>
          <Pivot aria-label="Count and Icon Pivot Example">
            <PivotItem headerText="My Files" itemCount={42} itemIcon="Emoji2">
              <Label styles={labelStyles}>Pivot #1</Label>
            </PivotItem>
            <PivotItem itemCount={23} itemIcon="Recent">
              <Label styles={labelStyles}>Pivot #2</Label>
            </PivotItem>
            <PivotItem itemCount={23} itemIcon="Recent">
              <Label styles={labelStyles}>Pivot #2</Label>
            </PivotItem>
            <PivotItem headerText="Teaching buble" itemIcon="Recent">
              <Label styles={labelStyles}>Teaching buble</Label>
              <button id="botonTeaching">botonTeaching</button>
              <TeachingBubble
                target={`#botonTeaching`}
                primaryButtonProps={examplePrimaryButtonProps}
                onDismiss={()=>{}}
                headline="Discover what’s trending around you"
              >
                Lorem ipsum dolor sit amet, consectetur adipisicing elit. Facere, nulla, ipsum? Molestiae quis aliquam magni
                harum non?
              </TeachingBubble>
            </PivotItem>
            <PivotItem headerText="Coachmark" itemIcon="Globe">
              <Label styles={labelStyles}>Coachmark</Label>
              <button id="botoncitoCoachmark">botoncitoCoachmark</button>
              <Coachmark
                target={botoncitoCoachmark}
                positioningContainerProps={{
                  directionalHint: DirectionalHint.leftBottomEdge,
                  doNotLayer: true,
                }}
                ariaAlertText="A coachmark has appeared"
                ariaDescribedBy="coachmark-desc1"
                ariaLabelledBy="coachmark-label1"
                ariaDescribedByText="Press enter or alt + C to open the coachmark notification"
                ariaLabelledByText="Coachmark notification"
              >
                <TeachingBubbleContent
                  headline="Example title"
                  hasCloseButton
                  closeButtonAriaLabel="Close"
                  primaryButtonProps={buttonProps}
                  secondaryButtonProps={buttonProps2}
                  onDismiss={()=>{}}
                  ariaDescribedBy="example-description1"
                  ariaLabelledBy="example-label1"
                >
                  Welcome to the land of coachmarks!
                </TeachingBubbleContent> 
              </Coachmark>

<Coachmark
  target={botoncitoCoachmark}
  positioningContainerProps={{
    directionalHint: DirectionalHint.rightBottomEdge,
    doNotLayer: true,
  }}
  ariaAlertText="A coachmark has appeared"
  ariaDescribedBy="coachmark-desc1"
  ariaLabelledBy="coachmark-label1"
  ariaDescribedByText="Press enter or alt + C to open the coachmark notification"
  ariaLabelledByText="Coachmark notification"
>
  <TeachingBubbleContent
    headline="Example title"
    hasCloseButton
    closeButtonAriaLabel="Close"
    primaryButtonProps={buttonProps}
    secondaryButtonProps={buttonProps2}
    onDismiss={()=>{}}
    ariaDescribedBy="example-description1"
    ariaLabelledBy="example-label1"
  >
    Welcome to the land of coachmarks!
  </TeachingBubbleContent>
</Coachmark>
            </PivotItem>
            <PivotItem headerText="Message bar" itemIcon="Ringer" itemCount={1}>
              <Label styles={labelStyles}>Message bar</Label>
              <Text block>
                By default, MessageBar renders its content within an internal live region after a short delay to help ensure
                it's announced by screen readers. You can disable this behavior (while still ensuring the message is read by
                screen readers) by setting the <code>delayedRender</code> prop to <code>false</code> and setting up the
                MessageBar in one of the following ways.
              </Text>

              <Toggle inlineLabel label="Show status example" checked={true} onChange={()=>{}} />

              <MessageBar
                  delayedRender={false}
                  role="none"
                  messageBarType={MessageBarType.error}
                >
                  This is a status message.
              </MessageBar>

              <MessageBar
                  delayedRender={false}
                  role="none"
                  messageBarType={MessageBarType.severeWarning}

                >
                  This is a status message.
              </MessageBar>

              <MessageBar
                  delayedRender={false}
                  role="none"
                  messageBarType={MessageBarType.warning}

                >
                  This is a status message.
              </MessageBar>

              <MessageBar
                  delayedRender={false}
                  role="none"
                  messageBarType={MessageBarType.info}

                >
                  This is a status message.
              </MessageBar>

              <MessageBar
                  delayedRender={false}
                  role="none"
                  truncated={true}
                  messageBarType={MessageBarType.blocked}
                  isMultiline={false}
                >
                  This is a status message.
                  This is a status message.
                  This is a status message.
                  This is a status message.
              </MessageBar>

              <MessageBar
                  delayedRender={true}
                  role="none"
                >
                  This is a status message.
              </MessageBar>

              <MessageBar
                  delayedRender={true}
                  role="alert"
                >
                  This is a status message.
              </MessageBar>

              <MessageBar
                  delayedRender={true}
                  role="status"
                >
                  This is a status message.
              </MessageBar>
            </PivotItem>
            <PivotItem headerText="Nav Component" itemIcon="Globe" itemCount={10} onRenderItemLink={_customRenderer}>
              <Label styles={labelStyles}>Nav component</Label>
              <Nav
                onLinkClick={_onLinkClick}
                selectedKey="key3"
                ariaLabel="Nav basic example"
                styles={navStyles}
                groups={navLinkGroups}
              />
            </PivotItem>
          </Pivot>
        </div>
        <div className="ms-welcome">
          <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
          <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
            <p className="ms-font-l">
              Modify the source files, then click <b>Run</b>.
            </p>
            <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
              Run
            </DefaultButton>
          </HeroList>
        </div>
      </>
    );
  }
}


function _onLinkClick(_: React.MouseEvent<HTMLElement>, item?: INavLink) {
  if (item && item.name === 'News') {
    alert('News link clicked');
  }
}


function _customRenderer(
  link?: IPivotItemProps,
  defaultRenderer?: (link?: IPivotItemProps) => JSX.Element | null,
): JSX.Element | null {
  if (!link || !defaultRenderer) {
    return null;
  }

  return (
    <span style={{ flex: '0 1 100%' }}>
      {defaultRenderer({ ...link, itemIcon: undefined })}
      <Icon iconName={link.itemIcon} style={{ color: 'red' }} />
    </span>
  );
}
