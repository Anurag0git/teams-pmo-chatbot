// Teams PMO Bot - Simplified Version for Internship
// Created by: Anurag Lashkare

const { ActivityHandler, MessageFactory, CardFactory } = require('botbuilder');
const restify = require('restify');

// Simple in-memory storage
const botData = {
    reminders: [],
    trainings: [],
    acknowledgments: {},
    teamStats: {
        totalReminders: 0,
        totalTrainings: 0,
        totalAcknowledgments: 0
    }
};

class PMOBot extends ActivityHandler {
    constructor() {
        super();
        
        // Handle incoming messages
        this.onMessage(async (context, next) => {
            const userMessage = context.activity.text.toLowerCase().trim();
            
            console.log(`Received message: ${userMessage}`);
            
            // Route messages to appropriate handlers
            if (userMessage.startsWith('/')) {
                await this.handleCommand(context, userMessage);
            } else if (userMessage.includes('help')) {
                await this.showHelp(context);
            } else if (userMessage.includes('remind')) {
                await this.handleReminderRequest(context);
            } else if (userMessage.includes('training')) {
                await this.handleTrainingRequest(context);
            } else if (userMessage.includes('status')) {
                await this.showDashboard(context);
            } else {
                await this.handleGeneralMessage(context);
            }
            
            await next();
        });

        // Welcome new users
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let member of membersAdded) {
                if (member.id !== context.activity.recipient.id) {
                    await this.sendWelcomeMessage(context);
                }
            }
            await next();
        });
    }

    async handleCommand(context, command) {
        const parts = command.split(' ');
        const cmd = parts[0];
        const args = parts.slice(1);

        switch (cmd) {
            case '/remind':
                await this.createReminder(context, args);
                break;
            case '/ack':
                await this.acknowledgeReminder(context, args);
                break;
            case '/training':
                await this.addTraining(context, args);
                break;
            case '/status':
                await this.showDashboard(context);
                break;
            case '/help':
                await this.showHelp(context);
                break;
            default:
                await context.sendActivity(`Unknown command: ${cmd}. Type '/help' for available commands.`);
        }
    }

    async createReminder(context, args) {
        if (args.length < 2) {
            await context.sendActivity('Usage: /remind [time] [message]\nExample: /remind 1h Submit weekly reports');
            return;
        }

        const timeStr = args[0];
        const message = args.slice(1).join(' ');
        const reminderTime = this.parseTime(timeStr);
        
        if (!reminderTime) {
            await context.sendActivity('Invalid time format. Use: 1h, 30m, 2d, etc.');
            return;
        }

        const reminder = {
            id: this.generateId(),
            message: message,
            scheduledTime: reminderTime,
            createdBy: context.activity.from.name,
            createdAt: new Date(),
            acknowledged: []
        };

        botData.reminders.push(reminder);
        botData.teamStats.totalReminders++;

        const card = this.createReminderCard(reminder);
        await context.sendActivity(MessageFactory.attachment(card));
        await context.sendActivity(`✅ Reminder created! ID: ${reminder.id}`);
    }

    async acknowledgeReminder(context, args) {
        if (args.length < 1) {
            await context.sendActivity('Usage: /ack [reminder-id]');
            return;
        }

        const reminderId = args[0];
        const reminder = botData.reminders.find(r => r.id === reminderId);
        
        if (!reminder) {
            await context.sendActivity('❌ Reminder not found. Use /status to see active reminders.');
            return;
        }

        const userName = context.activity.from.name;
        if (!reminder.acknowledged.includes(userName)) {
            reminder.acknowledged.push(userName);
            botData.teamStats.totalAcknowledgments++;
        }

        await context.sendActivity(`✅ ${userName} acknowledged: "${reminder.message}"`);
    }

    async addTraining(context, args) {
        if (args.length < 3) {
            await context.sendActivity('Usage: /training [title] [category] [due-date]\nExample: /training "Data Privacy" compliance 2024-07-15');
            return;
        }

        const training = {
            id: this.generateId(),
            title: args[0].replace(/"/g, ''),
            category: args[1],
            dueDate: args[2],
            createdBy: context.activity.from.name,
            createdAt: new Date(),
            completed: []
        };

        botData.trainings.push(training);
        botData.teamStats.totalTrainings++;

        const card = this.createTrainingCard(training);
        await context.sendActivity(MessageFactory.attachment(card));
        await context.sendActivity(`✅ Training added! ID: ${training.id}`);
    }

    async showDashboard(context) {
        const stats = this.calculateStats();
        
        const dashboardCard = {
            type: 'AdaptiveCard',
            version: '1.3',
            body: [
                {
                    type: 'TextBlock',
                    text: '📊 PMO Team Dashboard',
                    weight: 'Bolder',
                    size: 'Large',
                    color: 'Accent'
                },
                {
                    type: 'ColumnSet',
                    columns: [
                        {
                            type: 'Column',
                            width: 'stretch',
                            items: [
                                {
                                    type: 'TextBlock',
                                    text: '📅 Active Reminders',
                                    weight: 'Bolder',
                                    size: 'Medium'
                                },
                                {
                                    type: 'TextBlock',
                                    text: stats.activeReminders.toString(),
                                    size: 'ExtraLarge',
                                    color: 'Accent'
                                }
                            ]
                        },
                        {
                            type: 'Column',
                            width: 'stretch',
                            items: [
                                {
                                    type: 'TextBlock',
                                    text: '📚 Active Trainings',
                                    weight: 'Bolder',
                                    size: 'Medium'
                                },
                                {
                                    type: 'TextBlock',
                                    text: stats.activeTrainings.toString(),
                                    size: 'ExtraLarge',
                                    color: 'Good'
                                }
                            ]
                        }
                    ]
                },
                {
                    type: 'FactSet',
                    facts: [
                        { title: 'Total Acknowledgments', value: stats.totalAcknowledgments.toString() },
                        { title: 'Pending Items', value: stats.pendingItems.toString() },
                        { title: 'Team Efficiency', value: stats.efficiency + '%' }
                    ]
                }
            ]
        };

        await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(dashboardCard)));
        
        // Show recent reminders
        if (botData.reminders.length > 0) {
            let recentReminders = '📋 **Recent Reminders:**\n';
            botData.reminders.slice(-3).forEach(reminder => {
                recentReminders += `• ${reminder.message} (ID: ${reminder.id}) - ${reminder.acknowledged.length} acknowledged\n`;
            });
            await context.sendActivity(recentReminders);
        }
    }

    async showHelp(context) {
        const helpCard = {
            type: 'AdaptiveCard',
            version: '1.3',
            body: [
                {
                    type: 'TextBlock',
                    text: '🤖 PMO Assistant - Help Guide',
                    weight: 'Bolder',
                    size: 'Large',
                    color: 'Accent'
                },
                {
                    type: 'TextBlock',
                    text: '**Available Commands:**',
                    weight: 'Bolder',
                    size: 'Medium'
                },
                {
                    type: 'TextBlock',
                    text: '• `/remind [time] [message]` - Create reminder\n• `/ack [id]` - Acknowledge reminder\n• `/training [title] [category] [date]` - Add training\n• `/status` - Show dashboard\n• `/help` - Show this help',
                    wrap: true
                },
                {
                    type: 'TextBlock',
                    text: '**Examples:**',
                    weight: 'Bolder',
                    size: 'Medium'
                },
                {
                    type: 'TextBlock',
                    text: '• `/remind 2h Submit weekly reports`\n• `/ack abc123`\n• `/training "Data Security" compliance 2024-07-15`\n• `help` or `status` (natural language)',
                    wrap: true
                }
            ]
        };

        await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(helpCard)));
    }

    async sendWelcomeMessage(context) {
        const welcomeText = `👋 Welcome to PMO Assistant!\n\nI can help you with:\n• Setting up team reminders\n• Tracking acknowledgments\n• Managing training schedules\n• Monitoring team status\n\nType 'help' to get started!`;
        await context.sendActivity(MessageFactory.text(welcomeText));
    }

    async handleGeneralMessage(context) {
        const responses = [
            "I'm here to help with PMO tasks! Type 'help' to see what I can do.",
            "Need assistance? I can help with reminders, training, and team management.",
            "Try '/status' to see your team dashboard or '/help' for commands."
        ];
        
        const randomResponse = responses[Math.floor(Math.random() * responses.length)];
        await context.sendActivity(randomResponse);
    }

    async handleReminderRequest(context) {
        await context.sendActivity("To create a reminder, use: `/remind [time] [message]`\nExample: `/remind 1h Team meeting in conference room`");
    }

    async handleTrainingRequest(context) {
        await context.sendActivity("To add training, use: `/training [title] [category] [due-date]`\nExample: `/training 'Compliance Training' mandatory 2024-07-15`");
    }

    createReminderCard(reminder) {
        const card = {
            type: 'AdaptiveCard',
            version: '1.3',
            body: [
                {
                    type: 'TextBlock',
                    text: '📅 New Reminder',
                    weight: 'Bolder',
                    size: 'Medium',
                    color: 'Accent'
                },
                {
                    type: 'TextBlock',
                    text: reminder.message,
                    weight: 'Bolder',
                    wrap: true
                },
                {
                    type: 'FactSet',
                    facts: [
                        { title: 'Scheduled', value: reminder.scheduledTime.toLocaleString() },
                        { title: 'Created by', value: reminder.createdBy },
                        { title: 'ID', value: reminder.id }
                    ]
                }
            ]
        };
        
        return CardFactory.adaptiveCard(card);
    }

    createTrainingCard(training) {
        const card = {
            type: 'AdaptiveCard',
            version: '1.3',
            body: [
                {
                    type: 'TextBlock',
                    text: '📚 Training Added',
                    weight: 'Bolder',
                    size: 'Medium',
                    color: 'Good'
                },
                {
                    type: 'TextBlock',
                    text: training.title,
                    weight: 'Bolder',
                    wrap: true
                },
                {
                    type: 'FactSet',
                    facts: [
                        { title: 'Category', value: training.category },
                        { title: 'Due Date', value: training.dueDate },
                        { title: 'Created by', value: training.createdBy },
                        { title: 'ID', value: training.id }
                    ]
                }
            ]
        };
        
        return CardFactory.adaptiveCard(card);
    }

    parseTime(timeStr) {
        const match = timeStr.match(/^(\d+)([hmwd])$/);
        if (!match) return null;

        const value = parseInt(match[1]);
        const unit = match[2];
        const now = new Date();

        switch (unit) {
            case 'm': return new Date(now.getTime() + value * 60 * 1000);
            case 'h': return new Date(now.getTime() + value * 60 * 60 * 1000);
            case 'd': return new Date(now.getTime() + value * 24 * 60 * 60 * 1000);
            case 'w': return new Date(now.getTime() + value * 7 * 24 * 60 * 60 * 1000);
            default: return null;
        }
    }

    generateId() {
        return Math.random().toString(36).substr(2, 6);
    }

    calculateStats() {
        const totalAcknowledgments = botData.reminders.reduce((sum, r) => sum + r.acknowledged.length, 0);
        const pendingItems = botData.reminders.length + botData.trainings.length;
        const efficiency = botData.reminders.length > 0 ? Math.round((totalAcknowledgments / botData.reminders.length) * 100) : 100;

        return {
            activeReminders: botData.reminders.length,
            activeTrainings: botData.trainings.length,
            totalAcknowledgments: totalAcknowledgments,
            pendingItems: pendingItems,
            efficiency: efficiency
        };
    }
}

// Create HTTP server
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

// Create bot adapter
const { BotFrameworkAdapter } = require('botbuilder');
const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId || '',
    appPassword: process.env.MicrosoftAppPassword || ''
});

// Create bot instance
const bot = new PMOBot();

// Start server
const port = process.env.PORT || 3978;
server.listen(port, () => {
    console.log(`🤖 PMO Bot is running on port ${port}`);
    console.log(`📱 Bot endpoint: http://localhost:${port}/api/messages`);
    console.log(`🚀 Ready to help with PMO tasks!`);
});

// Handle incoming requests
server.post('/api/messages', async (req, res) => {
    await adapter.process(req, res, (context) => bot.run(context));
});

// Error handling
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] Error: ${error}`);
    await context.sendActivity('Sorry, I encountered an error. Please try again.');
};

// Graceful shutdown
process.on('SIGINT', () => {
    console.log('\n🛑 Shutting down PMO Bot...');
    server.close(() => {
        console.log('✅ Bot stopped successfully');
        process.exit(0);
    });
});

console.log('🔧 PMO Bot initialized successfully!');
console.log('📋 Features loaded: Reminders, Training, Acknowledgments, Dashboard');
console.log('💡 Tip: Use Bot Framework Emulator to test locally');