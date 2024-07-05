
import { ActivityHandler, MessageFactory, CardFactory } from 'botbuilder';
import { DialogSet, WaterfallDialog, WaterfallStepContext, DialogTurnStatus } from 'botbuilder-dialogs';

const DIALOG_ID = 'sendMessageDialog';

export class TeamsBot extends ActivityHandler {
    private dialogState: any;
    private dialogs: DialogSet;

    constructor(conversationState: any) {
        super();

        this.dialogState = conversationState.createProperty('dialogState');
        this.dialogs = new DialogSet(this.dialogState);

        // Add dialogs
        this.dialogs.add(new WaterfallDialog(DIALOG_ID, [
            this.showDialogStep.bind(this),
            this.sendMessageStep.bind(this)
        ]));

        // Handle message activity
        this.onMessage(async (context, next) => {
            const dialogContext = await this.dialogs.createContext(context);
            const result = await dialogContext.continueDialog();

            if (result.status === DialogTurnStatus.empty) {
                await dialogContext.beginDialog(DIALOG_ID);
            }

            await next();
        });
    }

    private async showDialogStep(step: WaterfallStepContext) {
        const card = CardFactory.heroCard(
            'Send a Message',
            undefined,
            [
                {
                    type: 'invoke',
                    title: 'Select User',
                    value: {
                        type: 'task/fetch'
                    }
                }
            ]
        );

        await step.context.sendActivity(MessageFactory.attachment(card));
        return await step.next();
    }

    private async sendMessageStep(step: WaterfallStepContext) {
        const user = step.result;
        await step.context.sendActivity(`Message sent to ${user.name}`);
        return await step.endDialog();
    }
}
