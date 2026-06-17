import type { IWordSession } from "../word/session.js"
import type { PositionMap } from "../word/position-map.js"
import type { SessionDirector } from "./session-director.js"

export interface ServerContext {
  session: IWordSession | null
  positionMap: PositionMap | null
  director: SessionDirector | null
}

export interface ReadyServerContext extends ServerContext {
  session: IWordSession
  positionMap: PositionMap
  director: SessionDirector
}
