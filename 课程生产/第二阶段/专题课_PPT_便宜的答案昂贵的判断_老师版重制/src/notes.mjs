export function teacherNotes({ minutes, objective, script, action, reaction, followup, caution, transition, source }) {
  return [
    `[时长] ${minutes}分钟`,
    `[目标] ${objective}`,
    `[老师话术] ${script}`,
    `[动作] ${action}`,
    `[预期反应] ${reaction}`,
    `[点评追问] ${followup}`,
    `[口径提醒] ${caution}`,
    `[转场] ${transition}`,
    `[讲义来源] ${source}`,
  ].join("\n");
}

export const REQUIRED_NOTE_FIELDS = [
  "[时长]", "[目标]", "[老师话术]", "[动作]", "[预期反应]",
  "[点评追问]", "[口径提醒]", "[转场]", "[讲义来源]",
];
