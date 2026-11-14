import QtQuick
import qs.Commons
import qs.Commons as Commons

import "Strings.js" as Strings

// The Accounts section of Settings: the connected accounts, each with its
// address and whether its sign-in still works, then a kind to pick and a
// button that connects one more. Connecting is a sign-in in a browser, so
// the button opens the setup in a terminal window with the right option.
Column {
  id: root

  property color foreground: Commons.Color.foreground
  property string fontFamily: Style.font.family
  property string language: "en"
  // [{ type: "google" | "outlook", number, address, calendars, ok, error }]
  property var accounts: []
  property bool loading: false
  // What the last Connect led to: "" | "opened" | an error to show as it is.
  property string notice: ""
  property string kind: "google"

  signal connectRequested(string kind)

  readonly property color faint: Qt.rgba(foreground.r, foreground.g, foreground.b, 0.50)
  readonly property var kindNames: ({ google: "Google", outlook: "Outlook / Microsoft 365" })
  // The first account of a kind has more to set up than the ones after it.
  readonly property bool first: accounts.filter(function(account) {
    return account.type === root.kind
  }).length === 0

  function tr(key, args) {
    return Strings.tr(root.language, key, args)
  }

  spacing: Style.space(8)

  SettingsCaption {
    foreground: root.foreground
    fontFamily: root.fontFamily
    text: root.tr("accounts.connected")
    note: root.loading ? root.tr("accounts.checking") : String(root.accounts.length)
  }

  Text {
    visible: !root.loading && root.accounts.length === 0
    width: parent.width
    text: root.tr("accounts.none")
    color: root.faint
    font.family: root.fontFamily
    font.pixelSize: Style.font.caption
    wrapMode: Text.WordWrap
  }

  Repeater {
    model: root.accounts

    Column {
      id: accountRow
      required property var modelData

      width: root.width
      leftPadding: Style.space(4)
      spacing: Style.space(1)

      Text {
        width: parent.width - parent.leftPadding
        // The address comes from the account's server.
        textFormat: Text.PlainText
        elide: Text.ElideRight
        text: (root.kindNames[accountRow.modelData.type] || "")
          + (accountRow.modelData.address ? " · " + accountRow.modelData.address : " " + accountRow.modelData.number)
        color: root.foreground
        font.family: root.fontFamily
        font.pixelSize: Style.font.bodySmall
      }

      Text {
        width: parent.width - parent.leftPadding
        textFormat: Text.PlainText
        wrapMode: Text.WordWrap
        text: accountRow.modelData.ok
          ? Strings.trn(root.language, "accounts.calendars", accountRow.modelData.calendars)
          : (accountRow.modelData.error || root.tr("accounts.broken"))
        color: accountRow.modelData.ok ? root.faint : Commons.Color.accent
        font.family: root.fontFamily
        font.pixelSize: Style.font.caption
      }
    }
  }

  Item { width: 1; height: Style.space(8) }

  SettingsCaption {
    foreground: root.foreground
    fontFamily: root.fontFamily
    text: root.tr("accounts.add")
  }

  ChoiceGroup {
    foreground: root.foreground
    fontFamily: root.fontFamily
    label: root.tr("accounts.kind")
    hint: root.tr(root.kind === "google"
      ? (root.first ? "accounts.googleFirst" : "accounts.googleMore")
      : (root.first ? "accounts.outlookFirst" : "accounts.outlookMore"))
    options: [
      { value: "google", label: root.kindNames.google },
      { value: "outlook", label: root.kindNames.outlook }
    ]
    value: root.kind
    onChosen: function(value) { root.kind = value }
  }

  AccentButton {
    foreground: root.foreground
    fontFamily: root.fontFamily
    iconText: "󰐕"
    text: root.tr("accounts.connect")
    onClicked: root.connectRequested(root.kind)
  }

  Text {
    visible: root.notice !== ""
    width: parent.width
    textFormat: Text.PlainText
    wrapMode: Text.WordWrap
    text: root.notice === "opened" ? root.tr("accounts.opened") : root.notice
    color: root.notice === "opened" ? root.faint : Commons.Color.accent
    font.family: root.fontFamily
    font.pixelSize: Style.font.caption
  }
}
