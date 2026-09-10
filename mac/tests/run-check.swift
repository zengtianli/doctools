import Foundation
@main struct RunCheck {
 @MainActor static func main() async throws {
  let vm = AppViewModel()
  await vm.loadOps()
  precondition(!vm.ops.isEmpty)
  vm.selectedOpID = vm.ops.first!.id
  vm.files = [InputFile(id: "/nonexistent-dockit-test.md")]
  vm.isRunning = true
  vm.statusText = "existing job"
  await vm.run()
  precondition(vm.isRunning && vm.statusText == "existing job", "must reject duplicate run")
  print("PASS duplicate run rejected by production AppViewModel")
 }
}
