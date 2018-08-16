class GChatMessage : Message {

    [GChatMessageType]$MessageType = [GChatMessageType]::MESSAGE

    GChatMessage(
        [string]$To,
        [string]$From,
        [string]$Body = [string]::Empty
    ) {
        $this.To = $To
        $this.From = $From
        $this.Body = $Body
    }
}
