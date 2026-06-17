import type { ICollectionProxy } from "./types.js"

export class CollectionProxy<T> implements ICollectionProxy<T> {
  constructor(
    private rawCollection: Record<string, unknown>,
    private itemFactory: (raw: Record<string, unknown>) => T,
  ) {}

  get count(): number {
    return this.rawCollection.Count as number
  }

  item(index: number): T {
    const raw = (this.rawCollection.Item as (i: number) => Record<string, unknown>)(index)
    return this.itemFactory(raw)
  }
}
