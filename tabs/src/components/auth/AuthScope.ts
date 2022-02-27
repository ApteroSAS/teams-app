import { MgtPersonCard } from '@microsoft/mgt';

export function getScopes():string[]{
    const neededScopes = MgtPersonCard.getScopes();
    return neededScopes
}