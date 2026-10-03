import Foundation

@main
enum PreferencesCheck {
    @MainActor static func main() throws {
        let domain = "dockit-fixture-" + UUID().uuidString
        let preferences = UserDefaults(suiteName: domain)!
        defer { preferences.removePersistentDomain(forName: domain) }
        let fixture = """
        [{"id":"fixtureA","targets":[{"id":"a1"},{"id":"a2"}],
          "options":[{"id":"fixtureConsent","type":"bool","group":"privacy","default":false}]},
         {"id":"fixtureB","targets":[{"id":"b1"},{"id":"b2"}]}]
        """
        let decoder = JSONDecoder()
        let operations = try decoder.decode([DocOp].self, from: Data(fixture.utf8))
        preferences.set("fixtureB", forKey: "dockit.lastOperation")
        preferences.set(["fixtureB": "b2"], forKey: "dockit.targetFormats")
        let vm = AppViewModel(preferences: preferences)
        vm.ops = operations
        vm.selectedOpID = "fixtureA"
        vm.files = [InputFile(id: "/fixture/local-document.docx")]
        vm.reloadPortablePreferences()
        precondition(vm.selectedOpID == "fixtureB" && vm.selectedTargetID == "b2")
        precondition(vm.files.map(\.id) == ["/fixture/local-document.docx"])

        vm.selectedOpID = "fixtureA"
        vm.onOpChanged()
        vm.selectedTargetID = "a2"
        precondition(preferences.string(forKey: "dockit.lastOperation") == "fixtureA")
        precondition((preferences.dictionary(forKey: "dockit.targetFormats") as? [String: String])?["fixtureA"] == "a2")
        vm.optionValues["fixtureConsent"] = true
        vm.optionPaths["fixtureTemplate"] = "/fixture/local-template.docx"
        vm.onOpChanged()
        precondition(vm.optionValues["fixtureConsent"] == false && vm.optionPaths.isEmpty)
        precondition(Set(preferences.dictionaryRepresentation().keys.filter { $0.hasPrefix("dockit.") })
                     == Set(AppViewModel.portablePreferenceKeys))

        vm.selectedOpID = "fixtureB"
        vm.onOpChanged()
        vm.selectedOpID = "fixtureA"
        vm.onOpChanged()
        precondition(vm.selectedTargetID == "a2")
        preferences.set("unknownOperation", forKey: "dockit.lastOperation")
        preferences.set(["fixtureA": "unknownFormat"], forKey: "dockit.targetFormats")
        vm.selectedTargetID = nil
        vm.reloadPortablePreferences()
        precondition(vm.selectedOpID == "fixtureA" && vm.selectedTargetID == "a1")

        vm.isRunning = true
        preferences.set("fixtureB", forKey: "dockit.lastOperation")
        preferences.set(["fixtureB": "b2"], forKey: "dockit.targetFormats")
        vm.reloadPortablePreferences()
        precondition(vm.selectedOpID == "fixtureA")
        vm.isRunning = false
        precondition(vm.selectedOpID == "fixtureB" && vm.selectedTargetID == "b2")
        print("PASS operation/format preferences, unknown values, consent reset, local files and task isolation")
    }
}
