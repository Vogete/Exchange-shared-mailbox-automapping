class SyncModel {
    [Collections.Generic.List[Object]]$ObjectsToAdd
    [Collections.Generic.List[Object]]$ObjectsToRemove

    SyncModel(){
        $this.ObjectsToAdd = [Collections.Generic.List[Object]]@()
        $this.ObjectsToRemove = [Collections.Generic.List[Object]]@()
    }
}
