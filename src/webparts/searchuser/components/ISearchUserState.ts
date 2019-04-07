import { IUserItem } from './IUserItem';

export default interface IGraphConsumerState {
    users: Array<IUserItem>;
    searchFor: string;
  }