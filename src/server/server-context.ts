import type { IWordSession } from "../word/session.js"
import type { PositionMap } from "../word/position-map.js"
import type { SessionDirector } from "./session-director.js"
import type { ILogger } from "../logger.js"

export interface ServerContext {
  session: IWordSession | null
  positionMap: PositionMap | null
  director: SessionDirector | null
  logger?: ILogger
  traceId?: string
}

export interface ReadyServerContext extends ServerContext {
  session: IWordSession
  positionMap: PositionMap
  director: SessionDirector
}
