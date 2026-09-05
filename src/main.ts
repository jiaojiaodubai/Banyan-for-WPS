import { loadConfig } from "./utils/config"
import { registerRibbonCallbacks } from "./modules/ribbon"

async function bootstrap() {
	await loadConfig()
	registerRibbonCallbacks()
}

void bootstrap()
