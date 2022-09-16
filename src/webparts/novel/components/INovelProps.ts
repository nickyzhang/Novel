export interface INovelProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  datalist: Array<IBpmTodoApply>
}

export interface IBpmTodoApply {
  applyId: number;
  workspace: string;
  applyNum: string;
  applyName: string;
  link: string;
}